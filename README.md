# Voice to Keyboard Converter

A Windows application that converts voice commands to keyboard inputs, allowing hands-free control of your computer.

## Features

- Converts voice commands to keyboard inputs
- Supports basic navigation keys (arrows, enter, space, etc.)
- Supports function keys (F1-F12)
- Supports keyboard shortcuts (Ctrl+C, Ctrl+V, Alt+Tab, etc.)
- Simple, easy-to-use interface
- Real-time status updates

## Supported Commands

### Basic Navigation
- "up", "down", "left", "right" - Arrow keys
- "enter" - Enter key
- "space" - Space key
- "backspace" - Backspace key
- "tab" - Tab key
- "escape" - Escape key
- "home" - Home key
- "end" - End key
- "page up" - Page Up key
- "page down" - Page Down key
- "delete" - Delete key
- "insert" - Insert key

### Function Keys
- "f1" through "f12" - Function keys

### Modifier Keys
- "control" - Ctrl key
- "shift" - Shift key
- "alt" - Alt key
- "windows" - Windows key

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

## Installation

1. Clone this repository or download the source code
2. Open the solution in Visual Studio
3. Build and run the application

## Usage

1. Launch the application
2. Click "Start Listening" button
3. Speak a command (e.g., "up", "control c", "enter")
4. Click "Stop Listening" to pause voice recognition

## Technical Details

This application uses:
- Windows Forms for the UI
- System.Speech for speech recognition
- Windows API (user32.dll) for simulating keyboard input through SendInput

## License

MIT