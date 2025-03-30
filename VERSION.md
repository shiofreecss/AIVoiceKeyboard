# AI Voice to Keyboard Converter - Version History

## v1.0.1 (Current)

### Improvements
- Added automatic pause after 1.5 seconds of silence in String Mode
- Enhanced noise filtering to remove various sound effect annotations:
  - Asterisk-enclosed noise like *cough*, *sad music*
  - Curly brace annotations like {sound effects}
  - Hash-surrounded noise like #background noise#
  - Tilde-enclosed items like ~background noise~
- Added filtering for common sound effect words (cough, sigh, laugh, music, etc.)
- Improved text cleanup after noise removal:
  - Better capitalization handling
  - Proper spacing around punctuation
  - Fixed standalone "i" to "I"
  - Balanced delimiters for quotes and parentheses
- Updated UI to reflect automatic silence detection

### Bug Fixes
- Fixed processing of longer audio segments with appropriate timeouts
- Improved sentence handling with proper punctuation

## v1.0.0 (Initial Release)

### New Features
- **Real-time Word Recognition**: Added real-time word-by-word processing for more responsive dictation
- **Improved Audio Capture**: Reduced audio capture time from 5s to 2.5s for faster recognition
- **Visual Speaking Indicator**: Added green "SPEAKING NOW - PLEASE TALK" indicator with flashing effect
- **Scrollable Text Display**: Replaced static label with scrollable text box for unlimited text length
- **Copy Functionality**: Added "Copy Text" button and right-click context menu with copy options
- **Noise Filtering**: Automatically removes noise annotations like (keyboard tapping), [buzzer], etc.
- **UI Improvements**: Updated color scheme and indicators for better visual feedback

### Core Features
- Converts voice commands to keyboard inputs
- Two recognition modes:
  - Command Mode: Execute keyboard commands and shortcuts
  - String Mode: Type spoken text directly with real-time word detection
- Support for basic navigation keys (arrows, enter, space, etc.)
- Support for function keys (F1-F12)
- Support for keyboard shortcuts (Ctrl+C, Ctrl+V, Alt+Tab, etc.)
- Advanced speech recognition using OpenAI's Whisper (offline, local processing)
- Visual speaking indicator that shows when to talk in String Mode
- Scrollable text display with copy functionality
- Basic noise annotation filtering
- Real-time status updates
