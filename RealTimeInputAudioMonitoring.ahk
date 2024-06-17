#Requires AutoHotkey v2.0 
Persistent

; Prepare GUI
myGui := Gui("-MinimizeBox")
myGui.AddText("x10 y10", "Amplitude (Loudness):")
myGui.AddProgress("w100 h180 x10 y30 c12921E Vertical BackgroundCFCFCF vAmpProgress")
myGui.OnEvent("Close", (*) => OnClose())

; Prepare ListView
listView := myGui.AddListView("w300 h200 x120 y10 Grid NoSort -Multi", ["Device Index", "Name"])
listView.OnEvent("Click", (LV, RowNumber) =>  recorder.SetDeviceIndex(RowNumber-1))

devices := AudioRecorder.GetInputDevices()
if devices.Length = 0{
    MsgBox("Error in retrieving the list of input devices")
    ExitApp
}

for i, deviceName in devices
    listView.Add(, A_Index-1, deviceName)

myGui.Show()

keepRecording := True

recorder := AudioRecorder(0)
While keepRecording
    recorder.Start(1, GetAmplitudeCallback)

OnClose(){
    keepRecording := False
    recorder.Stop()
    ExitApp
}

GetAmplitudeCallback(nRecordedSamples, nBytesPerSample, recordingBuffer) {
    static maxAvg := 1     ; Keep track of the max average so far to use it as an upper limit
    sum := 0               ; Stores the sum of the samples

    sampleDataType := nBytesPerSample = 2 ? "Short" : "UChar"
    Loop nRecordedSamples { 
        sample := NumGet(recordingBuffer, (A_Index - 1) * nBytesPerSample, sampleDataType)
        sum += Abs(sample) ; Abs() is used since any sample can be either negative or positive
    }

    avg := Ln((sum + 1) / (nRecordedSamples + 1) + 1)
    maxAvg := Max(maxAvg, avg)
    amplitude := Round((avg / maxAvg) * 100) ; Calculates the amplitude (loudness) as a percentage
    
    MyGui["AmpProgress"].Value := amplitude  ; Update the progressbar
}

class AudioRecorder {
    ;=========================================================================================
    ; Creates a new AudioRecorder instance
    ;=========================================================================================
    ; Parameters:
    ;   - deviceIndex (Integer): The index of the desired input device. 
    ;       Possible Values: (greater than 0)
    ;   - nSamplesPerSec (Integer): Number of samples to record per second
    ;       Possible Values: 8000, 11025, 22050, and 44100
    ;   - nBytesPerSample (Integer): Number of bytes per sample (Hz)
    ;       Possible Values: 1, 2
    ;   - nChannels (Integer): Number of channels in the waveform-audio data
    ;=========================================================================================
    __New(deviceIndex, nSamplesPerSec:=44100, nBytesPerSample:=2, nChannels:=1){
        this.deviceIndex := deviceIndex
        this.nSamplesPerSec := nSamplesPerSec
        this.nBytesPerSample := nBytesPerSample
        this.nChannels := nChannels                        ; Monaural data uses one channel and stereo data uses two channels.
        this.hwi := Buffer(A_PtrSize)                      ; Buffer to receive the handle
        this.WAVEFORMATEX := Buffer(18)                    ; Buffer to store the format of waveform-audio data
        this.WAVEHDR := Buffer(A_PtrSize * 4 + 16, 0)      ; Create WAVEHDR structure
        this.isPrepared := False                           ; A flag to monitor if Prepare() is executed
    }

    ;=========================================================================================
    ; (private) Prepares the environment for audio recording
    ;=========================================================================================
    ; Returns (Integer):
    ;   - 0: No errors
    ;   - 1: If an error has occurred in winmm\waveInOpen DLL call
    ;   - 2: If an error has occurred in winmm\waveInPrepareHeader DLL call
    ;=========================================================================================
    Prepare() {
        nBlockAlign := (this.nChannels * this.nBytesPerSample)
        nAvgBytesPerSec := this.nSamplesPerSec * nBlockAlign
       
        NumPut("UShort", 1, this.WAVEFORMATEX, 0)                 ; wFormatTag: (0x01 = WAVE_FORMAT_PCM)
        NumPut("UShort", this.nChannels, this.WAVEFORMATEX, 2)    ; nChannels: number of channels
        NumPut("UInt", this.nSamplesPerSec, this.WAVEFORMATEX, 4) ; Sample rate, in samples per second (hertz)
        NumPut("UInt", nAvgBytesPerSec, this.WAVEFORMATEX, 8)     ; Required average data-transfer rate, in bytes per second
        NumPut("UShort", nBlockAlign, this.WAVEFORMATEX, 12)      ; Block alignment, in bytes
        NumPut("UShort", this.nBytesPerSample * 8, this.WAVEFORMATEX, 14) ; Bits per sample for the wFormatTag format type
        NumPut("UShort", 0, this.WAVEFORMATEX, 16)                ; cbSize: Size, in bytes, of extra format information
   
        result := DllCall(
            "winmm\waveInOpen",                       ; Opens the given waveform-audio input device for recording.
            "Ptr",  this.hwi,                         ; Handle to the waveform input device
            "UInt", this.deviceIndex,                 ; Index of the input device to use
            "Ptr",  this.WAVEFORMATEX.ptr,            ; Pointer to WAVEFORMATEX struct
            "Ptr",  0,                                ; 0 Since we are using CALLBACK_NULL
            "Ptr",  0,                                ; Unused
            "UInt", 0                                 ; Flags for opening the device (CALLBACK_NULL = 0x00000000)
        )
    
        if (result)                                   ; Return if an error has occurred 
            return 1
       
        this.deviceHandle := NumGet(this.hwi, "Uint") ; Getting the device handle from the pointer "hwi"
    
        bufferSize := this.recordingDurationMS * this.nSamplesPerSec * this.nBytesPerSample // 1000       
        this.recordingBuffer := Buffer(bufferSize, 0) ; Ensure buffer is zeroed out
        
        ; Initialize the WAVEHDR structure
        NumPut("Ptr", this.recordingBuffer.ptr, this.WAVEHDR, 0)           ; Pointer to the waveform buffer (LPSTR lpData)
        NumPut("UInt", this.recordingBuffer.size, this.WAVEHDR, A_PtrSize) ; Length, in bytes, of the buffer (DWORD dwBufferLength)
    
        ; The waveInPrepareHeader function prepares a buffer for waveform-audio input
        result := DllCall("winmm\waveInPrepareHeader", "Ptr", this.deviceHandle, "Ptr", this.WAVEHDR.ptr, "UInt", this.WAVEHDR.size)
        if (result)
            return 2
        
        this.isPrepared := True
        return 0
    }
    
    ;=========================================================================================
    ; Starts recording input audio into a buffer
    ;=========================================================================================
    ; Parameters:
    ;   - recordingDurationMS (Integer): The length of each recording in milliseconds.
    ;       Possible Values: (greater than 0)
    ;   - callback (Function): A callback function used to process the recording buffer
    ;=========================================================================================
    ; Returns (Integer):
    ;   - 0: No errors
    ;   - 1: Error in Prepare() method
    ;   - 2: If an error has occurred in winmm\waveInAddBuffer DLL call
    ;   - 3: If an error has occurred in winmm\waveInStart DLL call
    ;   - 4: The recording buffer is empty
    ;=========================================================================================
    Start(recordingDurationMS, callback){
        this.recordingDurationMS := recordingDurationMS
        
        if not this.isPrepared
            if this.Prepare()
                return 1
    
        ; The waveInAddBuffer function sends an input buffer to the given waveform-audio input device
        result := DllCall("winmm\waveInAddBuffer", "Ptr", this.deviceHandle, "Ptr", this.WAVEHDR.ptr, "UInt", this.WAVEHDR.size)
        if (result)
            return 2

        ; The waveInStart function starts recording on the given waveform-audio input device
        result := DllCall("winmm\waveInStart", "Ptr", this.deviceHandle)
        if (result)
            return 3

        Loop  ; Waits for the buffer to be filled (WAVERR_STILLPLAYING = 33)
            dwFlagsDone := NumGet(this.WAVEHDR, A_PtrSize * 2 + 8, "UInt") & 0x00000001
        Until dwFlagsDone = 1

        dwBytesRecorded := NumGet(this.WAVEHDR, A_PtrSize + 4, "UInt")  ; Specifies the number of bytes recorded (DWORD dwBytesRecorded)

        if dwBytesRecorded = 0 ; Return if the device didn't record any data in the buffer
            return 4

        nRecordedSamples := dwBytesRecorded // this.nBytesPerSample

        callback(nRecordedSamples, this.nBytesPerSample, this.recordingBuffer)
    }

    ;=========================================================================================
    ; Stops the current ongoing recording
    ;=========================================================================================
    ; Returns (Integer):
    ;   - 0: No errors
    ;   - 1: If an error has occurred in winmm\waveInUnprepareHeader DLL call
    ;   - 2: If an error has occurred in winmm\waveInClose DLL call
    ;=========================================================================================
    Stop(){
        Loop  ; Waits for the buffer to be filled (WAVERR_STILLPLAYING = 33)
            ; The waveInUnprepareHeader function cleans up the preparation performed by the waveInPrepareHeader function
            result := DllCall("winmm\waveInUnprepareHeader", "Ptr", this.deviceHandle, "Ptr", this.WAVEHDR.ptr, "UInt", this.WAVEHDR.size)
        Until result != 33

        if result
            return 1

        Loop  ; Waits for the buffer to be filled (WAVERR_STILLPLAYING = 33)
            ; The waveInClose function closes the given waveform-audio input device
            result := DllCall("winmm\waveInClose", "Ptr", this.deviceHandle)
        Until result != 33  
        
        if result
            return 2

        return 0
    }

    SetDeviceIndex(deviceIndex){
        if deviceIndex >= 0 and deviceIndex < AudioRecorder.GetInputDevices().Length {
            this.deviceIndex := deviceIndex
            this.isPrepared := False
        }
    }

    ;=========================================================================================
    ; Retrives a list of input devices names  
    ;=========================================================================================
    ; Returns (Array[String]):
    ;   - Each element is 32 character long, represents an input device name.
    ;   - Each element's index represents deviceIndex which can be used when interacting 
    ;     with windows APIs.
    ;=========================================================================================
    static GetInputDevices(){
        devices := Array()                              ; Array to store devices' names
        numDevices := DllCall("winmm\waveInGetNumDevs") ; Get the number of audio input devices
        WAVEINCAPS := Buffer(80)                        ; Buffer to store device information

        Loop numDevices {
            DllCall("winmm\waveInGetDevCaps", "UInt", A_Index-1, "UInt", WAVEINCAPS.ptr, "UInt", WAVEINCAPS.size)
            devices.push(StrGet(WAVEINCAPS.ptr + 8, "UTF-16"))
        }

        return devices
    }
}
