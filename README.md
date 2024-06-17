# Real-Time AHK- v2 Input Audio Monitoring
Real-Time Input Audio Monitoring in AHK v2

![animation](https://github.com/balawi28/Real-Time-AHK-v2-Input-Audio-Monitoring/assets/41299807/bb0e7fd5-2120-4ae4-91ea-63c32b4c916c)

## Features
- Device Selection: Lists available input devices and allows the user to select one for monitoring.
- Amplitude Monitoring: Displays the current amplitude of the selected input device in a progress bar.
- Real-time Updates: Continuously updates the amplitude display in real-time.

## Implementation Details
This script utilizes the WinAPI Multimedia functions (winmm.dll) for audio recording and device management. The AudioRecorder class leverages functions such as waveInOpen, waveInPrepareHeader, and waveInStart to handle low-level audio input operations.

## Requirements
- AutoHotkey v2.0

## Contributing
Feel free to submit issues or pull requests if you find any bugs or have suggestions for improvements.

## License
This project is licensed under the WTFPL License.
