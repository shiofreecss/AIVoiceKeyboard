# AI Voice Keyboard - Version History

## v1.0.3 (Current) - 2025-04-10

### Improved Recording State Management
- Fixed issue with recording continuing to process after stopping
- Added proper cancellation token handling throughout recording pipeline
- Enhanced synchronization between main window and floating overlay button
- Improved UI state consistency when toggling recording

### Bug Fixes
- Fixed inconsistent button state in overlay when enabling floating button during active recording
- Added robust error handling in audio processing pipeline
- Improved thread safety throughout the application
- Enhanced UI refresh mechanisms to ensure visual state always matches actual state

## v1.0.2 - 2025-03-30

### Version Configuration System
- Added version.cfg file to manage application version information
- Added floating overlay button for easy access
- Application now reads version, build number, and release date from external configuration
- Improved About dialog shows detailed version and build information
- Simplified version management for future updates

### Bug Fixes
- Fixed occasional Whisper model initialization issues
- Improved error handling for audio capture failures
- Fixed threading issues that could cause UI freezes
- Enhanced recovery from recognition errors

## v1.0.1 - 2025-03-15

### Enhanced Noise Filtering
- Comprehensive filtering of noise annotations in various formats:
  - Asterisk-enclosed noise like *cough*, *sad music*
  - Parentheses, brackets, curly braces, and other delimiters
  - Common sound effect words with surrounding punctuation
- Intelligent text cleanup with proper capitalization and punctuation
- Better handling of sentences with balanced delimiters

### Automatic Speech Pause Detection
- Recording automatically stops after 1.5 seconds of silence
- No need to manually stop recording between phrases
- More natural speech flow with automatic pausing
- Visual indicator shows "WILL PAUSE AFTER 1.5s SILENCE"

### Improved Text Processing
- Better sentence handling with proper punctuation
- Fixed standalone "i" to proper "I" capitalization
- Appropriate spacing around punctuation marks
- Balanced delimiters for quotes and parentheses

### UI Improvements
- Cleaner status updates with color-coded messages
- Improved visual feedback during recording
- Enhanced error messaging with troubleshooting guidance

## v1.0.0 - 2025-03-01

### Initial Release
- Two recognition modes: Command Mode and String Mode
- AI-powered speech recognition using Whisper
- Windows Speech Recognition as fallback
- Real-time word-by-word recognition
- Visual speaking indicator
- Scrollable text display
- Copy and clear functionality
- Noise filtering
- Floating overlay option
- System tray integration

