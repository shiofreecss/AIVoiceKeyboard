# AI Voice Keyboard (v1.0.2)

A Windows application that converts voice commands to keyboard inputs, allowing hands-free control of your computer. The application offers both Windows Speech Recognition and OpenAI Whisper for more accurate speech-to-text capabilities.

Developed by [ShioDev](https://hello.shiodev.com) | Powered by [BeaverFoundation](https://beaver.foundation)

## Features

- Converts voice commands to keyboard inputs
- Real-time word-by-word speech recognition 
- Automatic pause after 1.5 seconds of silence in String Mode
- Enhanced noise filtering for cleaner transcription
- Supports basic navigation keys (arrows, enter, space, etc.)
- Supports function keys (F1-F12)
- Supports keyboard shortcuts (Ctrl+C, Ctrl+V, Alt+Tab, etc.)
- Two recognition modes:
  - **Command Mode**: Execute keyboard commands and shortcuts
  - **String Mode**: Type spoken text directly with real-time word detection
- Advanced speech recognition using OpenAI's Whisper (offline, local processing)
- Visual speaking indicator that shows when to talk in String Mode
- Scrollable text display with copy functionality
- Simple, easy-to-use interface
- Real-time status updates

## Supported Commands

### Mode Switching
- "command mode" or "command option" - Switch to command mode
- "string mode" or "string option" - Switch to string mode

### Basic Navigation
- "up", "down", "left", "right" - Arrow keys
- "enter" or "return" - Enter key
- "space" or "space bar" - Space key
- "backspace" or "back space" - Backspace key
- "tab" or "tab key" - Tab key
- "escape", "escape key", or "esc" - Escape key
- "home" or "top" - Home key
- "end" or "bottom" - End key
- "page up" - Page Up key
- "page down" - Page Down key
- "delete" or "delete key" - Delete key

### Function Keys
- "f1" through "f12" - Function keys

### Modifier Keys
- "control" or "ctrl" - Ctrl key
- "shift" - Shift key
- "alt" - Alt key
- "windows" or "win" - Windows key

### Numbers
- "zero" through "nine" - Number keys 0-9

### Punctuation
- "comma" - ,
- "period" - .
- "slash" - /
- "semicolon" - ;
- "quote" - '
- "bracket left" - [
- "bracket right" - ]
- "backslash" - \
- "minus" - -
- "equals" - =

### Common Keyboard Shortcuts
- "control c" - Copy
- "control v" - Paste
- "control x" - Cut
- "control z" - Undo
- "control y" - Redo
- "control s" - Save
- "control f" - Find
- "control a" - Select all
- "control p" - Print
- "control n" - New
- "control o" - Open
- "control w" - Close

### Windows Shortcuts
- "alt tab" - Switch between windows
- "alt f4" - Close window
- "windows d" - Show desktop
- "windows e" - Open File Explorer
- "windows r" - Open Run dialog

### Text Navigation
- "control left" - Move one word left
- "control right" - Move one word right
- "control home" - Move to beginning of document
- "control end" - Move to end of document

### Text Selection
- "shift left" - Select one character left
- "shift right" - Select one character right
- "shift up" - Select one line up
- "shift down" - Select one line down
- "control shift left" - Select one word left
- "control shift right" - Select one word right
- "shift delete" - Permanent delete (bypass Recycle Bin)

## Requirements

- Windows 10 or later
- .NET 8.0 or later
- Microphone
- 1.5+ GB free disk space (for Whisper medium model)

## Installation

1. Clone this repository or download the source code
2. Open the solution in Visual Studio
3. Build and run the application

### Setting Up Whisper

The application uses OpenAI's Whisper for high-quality speech recognition in String Mode. The first time you use Whisper:

1. Download the Whisper model file from [Hugging Face](https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-medium.bin)
2. Place the downloaded `ggml-medium.bin` file in the application's executable directory (typically `VoiceToKeyboard/bin/Debug/net8.0-windows/`)
3. The medium model (1.5GB) is set as default for best accuracy. You can also use smaller models like `ggml-tiny.bin` (75MB) or `ggml-small.bin` (450MB) by changing the source code if needed.

## Usage

1. Launch the application
2. Select the desired input mode:
   - **Command Mode**: For executing keyboard commands
   - **String Mode**: For typing spoken text
3. Check "Use AI Recognition" for more accurate speech recognition (requires model file)
4. Click "Start Listening" button
5. In String Mode, a green "SPEAK NOW - WILL PAUSE AFTER 1.5s SILENCE" indicator will appear when you can speak
6. Speak your commands or text - recording will automatically pause after 1.5 seconds of silence
7. The recognized speech will be displayed in the text area and typed automatically in real-time
8. Use the scrollbar to view longer text and the "Copy Text" button to copy to clipboard
9. Click "Stop Listening" when you're done with voice recognition

## New Features in v1.0.2 (Current)

### Version Configuration System
- Added a version.cfg file to manage application version information
- Application now reads version, build number, and release date from external configuration
- Improved About dialog shows detailed version and build information
- Simplified version management for future updates

## Features from v1.0.1

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

## Features from v1.0.0

### Real-time Word-by-Word Recognition
- Words appear as you speak them, without waiting for full sentences
- Faster response time with reduced audio capture length (2.5s vs 5s)
- Immediate visual feedback as each word is recognized
- Significantly reduced latency between speaking and text appearing

### Visual Speaking Indicator
- A green "SPEAKING NOW - PLEASE TALK" indicator appears when the system is actively recording
- Helps you know exactly when to speak for optimal recognition
- Accompanied by a green dot indicator in the status area

### Scrollable Text Display
- All recognized text is shown in a scrollable text box
- Automatically scrolls to show the latest text
- Allows reviewing longer dictation sessions

### Copy and Clear Functionality
- "Copy Text" button to copy all recognized text to clipboard
- Right-click context menu with Copy, Select All, and Clear options
- "Clear Text" button to start fresh

### Noise Filtering
- Automatically filters out noise annotations like (keyboard tapping), [buzzer], etc.
- Provides cleaner transcription without distracting annotations

## Tips for Better Recognition

- Speak clearly and at a moderate pace
- Use a good quality microphone in a quiet environment
- For Command Mode, use exact command phrases
- In String Mode, speak one word at a time for best real-time results
- Wait for the "SPEAKING NOW" indicator to appear before talking
- The medium model provides better accuracy but requires more system resources
- For faster word recognition with slightly lower accuracy, try the small model

## Version History

For a detailed changelog of all versions, see [VERSION.md](VERSION.md).

## Technical Details

This application uses:
- Windows Forms for the UI
- System.Speech for Windows built-in speech recognition
- OpenAI's Whisper (via Whisper.net) for enhanced speech recognition
- NAudio for audio capture
- Windows API (user32.dll) for simulating keyboard input through SendInput

## License

GNU General Public License v3.0 (GPL-3.0)

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program. If not, see <https://www.gnu.org/licenses/>.
