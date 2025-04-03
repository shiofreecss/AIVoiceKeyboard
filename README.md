# AI Voice Keyboard

A Windows application that converts voice to keyboard input using speech recognition technology.

![AI Voice Keyboard](VoiceToKeyboard/icon/ai-voice-keyboard.png)

## Features

- **Two Recognition Modes:**
  - **Command Mode**: Execute keyboard commands like "up", "down", "control c", etc.
  - **String Mode**: Type text as you speak naturally

- **AI-Powered Recognition:**
  - Uses Whisper.net for accurate speech recognition
  - Falls back to Windows Speech Recognition when needed

- **Convenient Interface:**
  - Floating overlay button for easy access
  - System tray integration
  - Minimalist design

## Getting Started

### Prerequisites

- Windows 10 or later
- .NET 8.0 Runtime
- Microphone for voice input

### Installation

1. Download the latest release from the [Releases](https://github.com/ShioDev/AIVoiceKeyboard/releases) page
2. Extract the ZIP file to a location of your choice
3. Run `VoiceToKeyboard.exe`

### First Run

On first launch, AI Voice Keyboard will need to download the Whisper model. If it doesn't download automatically:

1. Download the model manually from: https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-medium.bin
2. Place the model file in the same directory as the application executable

## Usage

### Basic Controls

- **Start/Stop Listening**: Click the "Start Listening" button or use the floating button
- **Switch Modes**: Select Command or String mode using the radio buttons
- **Floating Button**: Enable/disable the floating overlay for easy access

### Voice Commands

#### Command Mode

Say commands like:
- Navigation: "up", "down", "left", "right", "page up", "page down", "home", "end"
- Function keys: "f1", "f2", etc.
- Common shortcuts: "control c" (copy), "control v" (paste), "control z" (undo)
- Switch modes: "command option" or "string option"

#### String Mode

- Simply speak naturally and the system will type what you say
- Speech will automatically pause after 1.5 seconds of silence
- To stop recognition, click the recording button again

## Troubleshooting

- **No Speech Recognition**: Ensure your microphone is properly connected and set as the default recording device
- **Poor Recognition**: Speak clearly and at a moderate pace
- **Missing Whisper Model**: Download the model file manually as described in the First Run section

## License

This project is licensed under the GPL-3.0 License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Whisper.net](https://github.com/Whisper-net/Whisper.net) for AI speech recognition
- [NAudio](https://github.com/naudio/NAudio) for audio processing
- Developed by [ShioDev](https://hello.shiodev.com)
- Powered by [Beaver Foundation](https://beaver.foundation)

## Version

See [VERSION.md](VERSION.md) for version history and release notes.
