package main

import (
	"context"
	"encoding/binary"
	"fmt"
	"math/rand"
	"os"
	"os/signal"
	"time"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/go-vgo/robotgo"
	"github.com/moutend/go-wca/pkg/wca"
)

func main() {
	signalChan := make(chan os.Signal, 1)
	signal.Notify(signalChan, os.Interrupt)
	ctx, cancel := context.WithCancel(context.Background())
	successes := 0
	timeouts := 0

	go func() {
		<-signalChan
		fmt.Println("Interrupted by SIGINT")
		fmt.Println("Successes:", successes)
		fmt.Println("Timeouts:", timeouts)
		cancel()
	}()

	var err error
	if err = ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED); err != nil {
		return
	}
	defer ole.CoUninitialize()

	var de *wca.IMMDeviceEnumerator
	if err = wca.CoCreateInstance(wca.CLSID_MMDeviceEnumerator, 0, wca.CLSCTX_ALL, wca.IID_IMMDeviceEnumerator, &de); err != nil {
		return
	}
	defer de.Release()

	var mmd *wca.IMMDevice
	if err = de.GetDefaultAudioEndpoint(wca.ERender, wca.EConsole, &mmd); err != nil {
		return
	}
	defer mmd.Release()

	var ps *wca.IPropertyStore
	if err = mmd.OpenPropertyStore(wca.STGM_READ, &ps); err != nil {
		return
	}
	defer ps.Release()

	var pv wca.PROPVARIANT
	if err = ps.GetValue(&wca.PKEY_Device_FriendlyName, &pv); err != nil {
		return
	}
	fmt.Printf("Capturing audio from: %s\n", pv.String())

	var ac *wca.IAudioClient
	if err = mmd.Activate(wca.IID_IAudioClient, wca.CLSCTX_ALL, nil, &ac); err != nil {
		return
	}
	defer ac.Release()

	var wfx *wca.WAVEFORMATEX
	if err = ac.GetMixFormat(&wfx); err != nil {
		return
	}
	defer ole.CoTaskMemFree(uintptr(unsafe.Pointer(wfx)))

	wfx.WFormatTag = 1
	wfx.NBlockAlign = (wfx.WBitsPerSample / 8) * wfx.NChannels // 16 bit stereo is 32bit (4 byte) per sample
	wfx.NAvgBytesPerSec = wfx.NSamplesPerSec * uint32(wfx.NBlockAlign)
	wfx.CbSize = 0

	fmt.Println("--------")
	fmt.Printf("Format: PCM %d bit signed integer\n", wfx.WBitsPerSample)
	fmt.Printf("Rate: %d Hz\n", wfx.NSamplesPerSec)
	fmt.Printf("Channels: %d\n", wfx.NChannels)
	fmt.Println("--------")

	var defaultPeriod wca.REFERENCE_TIME
	var minimumPeriod wca.REFERENCE_TIME
	var latency time.Duration
	if err = ac.GetDevicePeriod(&defaultPeriod, &minimumPeriod); err != nil {
		return
	}
	latency = time.Duration(int(defaultPeriod) * 100)

	fmt.Println("Default period: ", defaultPeriod)
	fmt.Println("Minimum period: ", minimumPeriod)
	fmt.Println("Latency: ", latency)

	if err = ac.Initialize(wca.AUDCLNT_SHAREMODE_SHARED, wca.AUDCLNT_STREAMFLAGS_LOOPBACK, wca.REFERENCE_TIME(400*10000), 0, wfx, nil); err != nil {
		return
	}

	var bufferFrameSize uint32
	if err = ac.GetBufferSize(&bufferFrameSize); err != nil {
		return
	}
	fmt.Printf("Allocated buffer size: %d\n", bufferFrameSize)

	var acc *wca.IAudioCaptureClient
	if err = ac.GetService(wca.IID_IAudioCaptureClient, &acc); err != nil {
		return
	}
	defer acc.Release()

	if err = ac.Start(); err != nil {
		return
	}
	fmt.Println("Start loopback capturing with shared timer driven mode")
	fmt.Println("Press Ctrl-C to stop capturing")
	time.Sleep(latency)

	var offset int
	var isCapturing bool = true
	var data *byte
	var b *byte
	var availableFrameSize uint32
	var flags uint32
	var devicePosition uint64
	var qcpPosition uint64
	max := int64(0)
	var lastCast time.Time
	var lastLoot time.Time
	rand.Seed(time.Now().UnixNano())
	randMin := 2000
	randMax := 4000
	consecutiveTimeouts := 0

	for {
		if !isCapturing {
			break
		}
		select {
		case <-ctx.Done():
			isCapturing = false
		default:
			if err = acc.GetBuffer(&data, &availableFrameSize, &flags, &devicePosition, &qcpPosition); err != nil {
				continue
			}
			if availableFrameSize == 0 {
				continue
			}

			start := unsafe.Pointer(data)
			lim := int(availableFrameSize) * int(wfx.NBlockAlign)
			buf := make([]byte, lim)

			for n := 0; n < lim; n++ {
				b = (*byte)(unsafe.Pointer(uintptr(start) + uintptr(n)))
				buf[n] = *b
				if n%4 == 3 {
					val, _ := binary.Varint(buf[n-3 : n+1])
					if val > max {
						max = val
						// fmt.Println("new max", max)
						if max > 8100000 && lastLoot.IsZero() && time.Since(lastCast).Seconds() > 2 {
							fmt.Println("click", max)
							fmt.Println("max", max)
							robotgo.Click("center")
							lastLoot = time.Now()
							successes++
							consecutiveTimeouts = 0
						}
					}

					if !lastCast.IsZero() && time.Since(lastCast).Seconds() > 25 {
						if consecutiveTimeouts >= 10 {
							fmt.Println("too many timeouts")
							lastLoot = time.Time{}
							lastCast = time.Time{}
						} else {
							fmt.Println("too long", max)
							lastLoot = lastCast
							timeouts++
							consecutiveTimeouts++
						}
					}

					if !lastLoot.IsZero() && time.Since(lastLoot).Milliseconds() > int64(rand.Intn(randMax-randMin+1)+randMin) {
						lastLoot = time.Time{}
						max = val
						fmt.Println("reset", val)
						robotgo.Click("center")
						lastCast = time.Now()
					}
				}
			}

			offset += lim

			if err = acc.ReleaseBuffer(availableFrameSize); err != nil {
				return
			}

			time.Sleep(latency / 2)
		}
	}

	fmt.Println("Stop capturing")

	if err = ac.Stop(); err != nil {
		return
	}
}
