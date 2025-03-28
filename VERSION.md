# AI Voice to Keyboard - Version History

## Version 1.0.2 (Current) - March 28, 2025

### New Features
- **Real-time Word Recognition**: Added real-time word-by-word processing for more responsive dictation
- **Improved Audio Capture**: Reduced audio capture time from 5s to 2.5s for faster recognition
- **Visual Speaking Indicator**: Added green "SPEAKING NOW - PLEASE TALK" indicator with flashing effect
- **Scrollable Text Display**: Replaced static label with scrollable text box for unlimited text length
- **Copy Functionality**: Added "Copy Text" button and right-click context menu with copy options
- **Noise Filtering**: Automatically removes noise annotations like (keyboard tapping), [buzzer], etc.
- **UI Improvements**: Updated color scheme and indicators for better visual feedback

### Technical Improvements
- Switched to medium Whisper model (1.5GB) for better accuracy
- Reduced minimum audio length from 1000ms to 500ms for faster word detection
- Implemented 1.5 second timeout for Whisper processing to maintain responsiveness
- Added word-by-word text output with feedback after each recognized word
- Improved error handling and recovery for more stable operation

### UI Changes
- Added right-click context menu with Copy, Select All, and Clear options
- Updated application title to "AI Voice to Keyboard v1.0.2"
- Improved status messages with clear real-time feedback
- Added visual indicator showing when recording is active
- Adjusted layout for better usability

## Version 1.0.1 - March 20, 2025

Initial release with basic functionality.
