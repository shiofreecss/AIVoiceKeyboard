using System;
using System.Collections.Generic;
using System.Speech.Recognition;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace VoiceToKeyboard;

public partial class Form1 : Form
{
    private SpeechRecognitionEngine? recognizer;
    private bool isListening = false;
    
    // Import SendInput from user32.dll
    [DllImport("user32.dll")]
    static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    // Structures for SendInput
    [StructLayout(LayoutKind.Sequential)]
    public struct INPUT
    {
        public uint type;
        public INPUTUNION union;
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct INPUTUNION
    {
        [FieldOffset(0)]
        public MOUSEINPUT mouseInput;
        [FieldOffset(0)]
        public KEYBDINPUT keyboardInput;
        [FieldOffset(0)]
        public HARDWAREINPUT hardwareInput;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct KEYBDINPUT
    {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct HARDWAREINPUT
    {
        public uint uMsg;
        public ushort wParamL;
        public ushort wParamH;
    }

    // Constants for SendInput
    const uint INPUT_KEYBOARD = 1;
    const uint KEYEVENTF_KEYDOWN = 0x0000;
    const uint KEYEVENTF_KEYUP = 0x0002;

    // Dictionary to map voice commands to key codes
    private Dictionary<string, ushort> commandToKeyMap = new Dictionary<string, ushort>
    {
        // Basic navigation
        { "up", 0x26 },           // Up arrow
        { "down", 0x28 },         // Down arrow
        { "left", 0x25 },         // Left arrow
        { "right", 0x27 },        // Right arrow
        { "enter", 0x0D },        // Enter
        { "space", 0x20 },        // Space
        { "backspace", 0x08 },    // Backspace
        { "tab", 0x09 },          // Tab
        { "escape", 0x1B },       // Escape
        { "home", 0x24 },         // Home
        { "end", 0x23 },          // End
        { "page up", 0x21 },      // Page Up
        { "page down", 0x22 },    // Page Down
        { "delete", 0x2E },       // Delete
        { "insert", 0x2D },       // Insert
        
        // Function keys
        { "f1", 0x70 },           // F1
        { "f2", 0x71 },           // F2
        { "f3", 0x72 },           // F3
        { "f4", 0x73 },           // F4
        { "f5", 0x74 },           // F5
        { "f6", 0x75 },           // F6
        { "f7", 0x76 },           // F7
        { "f8", 0x77 },           // F8
        { "f9", 0x78 },           // F9
        { "f10", 0x79 },          // F10
        { "f11", 0x7A },          // F11
        { "f12", 0x7B },          // F12
        
        // Common keyboard shortcuts
        { "copy", 0x43 },         // C (used with Ctrl)
        { "paste", 0x56 },        // V (used with Ctrl)
        { "cut", 0x58 },          // X (used with Ctrl)
        { "undo", 0x5A },         // Z (used with Ctrl)
        { "redo", 0x59 },         // Y (used with Ctrl)
        { "save", 0x53 },         // S (used with Ctrl)
        { "find", 0x46 },         // F (used with Ctrl)
        { "select all", 0x41 },   // A (used with Ctrl)
        { "print", 0x50 },        // P (used with Ctrl)
        { "new", 0x4E },          // N (used with Ctrl)
        { "open", 0x4F },         // O (used with Ctrl)
        { "close", 0x57 },        // W (used with Ctrl)
        
        // Modifier keys (handled separately in recognition)
        { "control", 0x11 },      // Ctrl
        { "shift", 0x10 },        // Shift
        { "alt", 0x12 },          // Alt
        { "windows", 0x5B },      // Windows key
        
        // Number keys
        { "zero", 0x30 },         // 0
        { "one", 0x31 },          // 1
        { "two", 0x32 },          // 2
        { "three", 0x33 },        // 3
        { "four", 0x34 },         // 4
        { "five", 0x35 },         // 5
        { "six", 0x36 },          // 6
        { "seven", 0x37 },        // 7
        { "eight", 0x38 },        // 8
        { "nine", 0x39 },         // 9
        
        // Punctuation
        { "comma", 0xBC },        // ,
        { "period", 0xBE },       // .
        { "slash", 0xBF },        // /
        { "semicolon", 0xBA },    // ;
        { "quote", 0xDE },        // '
        { "bracket left", 0xDB }, // [
        { "bracket right", 0xDD },// ]
        { "backslash", 0xDC },    // \
        { "minus", 0xBD },        // -
        { "equals", 0xBB },       // =
    };

    public Form1()
    {
        InitializeComponent();
        InitializeSpeechRecognition();
        InitializeUI();
    }

    private void InitializeUI()
    {
        // Change form title
        this.Text = "Voice to Keyboard";
        
        // Create a toggle button for starting/stopping recognition
        Button toggleButton = new Button
        {
            Text = "Start Listening",
            Size = new System.Drawing.Size(150, 50),
            Location = new System.Drawing.Point(325, 180),
            Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular)
        };
        toggleButton.Click += (sender, e) => ToggleRecognition(toggleButton);
        
        // Create a status label
        Label statusLabel = new Label
        {
            Text = "Status: Not Listening",
            AutoSize = true,
            Location = new System.Drawing.Point(330, 250),
            Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular)
        };
        this.Tag = statusLabel; // Store reference to update later
        
        // Create instructions label
        Label instructionsLabel = new Label
        {
            Text = "Say commands like \"up\", \"down\", \"enter\", \"control c\" for copy, etc.",
            AutoSize = true,
            Location = new System.Drawing.Point(150, 100),
            Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular)
        };
        
        // Add controls to form
        this.Controls.Add(toggleButton);
        this.Controls.Add(statusLabel);
        this.Controls.Add(instructionsLabel);
    }

    private void InitializeSpeechRecognition()
    {
        try
        {
            // Create the speech recognition engine
            recognizer = new SpeechRecognitionEngine();
            
            // Create grammar for commands
            Choices commands = new Choices();
            
            // Add all command keys 
            foreach (var command in commandToKeyMap.Keys)
            {
                commands.Add(command);
            }
            
            // Add combination commands
            commands.Add("control c");
            commands.Add("control v");
            commands.Add("control x");
            commands.Add("control z");
            commands.Add("control y");  // Redo
            commands.Add("control s");
            commands.Add("control f");  // Find
            commands.Add("control a");  // Select all
            commands.Add("control p");  // Print
            commands.Add("control n");  // New
            commands.Add("control o");  // Open
            commands.Add("control w");  // Close
            commands.Add("alt tab");
            commands.Add("alt f4");
            commands.Add("windows d");  // Show desktop
            commands.Add("windows e");  // Open explorer
            commands.Add("windows r");  // Run dialog
            commands.Add("shift delete"); // Permanent delete
            commands.Add("control home"); // Beginning of document
            commands.Add("control end");  // End of document
            commands.Add("shift left");   // Select left
            commands.Add("shift right");  // Select right
            commands.Add("shift up");     // Select up
            commands.Add("shift down");   // Select down
            commands.Add("control shift left");  // Select word left
            commands.Add("control shift right"); // Select word right
            commands.Add("control left");  // Word left
            commands.Add("control right"); // Word right
            
            // Create grammar with commands
            GrammarBuilder grammarBuilder = new GrammarBuilder();
            grammarBuilder.Append(commands);
            Grammar grammar = new Grammar(grammarBuilder);
            
            recognizer.LoadGrammar(grammar);
            recognizer.SpeechRecognized += Recognizer_SpeechRecognized;
            
            // Set input to default audio device
            recognizer.SetInputToDefaultAudioDevice();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Speech recognition initialization error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void Recognizer_SpeechRecognized(object? sender, SpeechRecognizedEventArgs e)
    {
        string recognizedText = e.Result.Text.ToLower();
        
        // Update status with recognized command
        if (this.Tag is Label statusLabel)
        {
            statusLabel.Text = $"Recognized: {recognizedText}";
            statusLabel.Refresh();
        }
        
        // Handle combined commands
        if (recognizedText.Contains(" "))
        {
            string[] parts = recognizedText.Split(' ');
            if (parts.Length == 2)
            {
                // Handle modifier + key combinations
                if (commandToKeyMap.TryGetValue(parts[0], out ushort modifierKey) && 
                    commandToKeyMap.TryGetValue(parts[1], out ushort actionKey))
                {
                    // Send modifier key down
                    SendKeyDown(modifierKey);
                    
                    // Send the action key press
                    SendKeyDown(actionKey);
                    SendKeyUp(actionKey);
                    
                    // Send modifier key up
                    SendKeyUp(modifierKey);
                }
            }
            else if (parts.Length == 3)
            {
                // Handle two modifiers + key combinations (e.g., "control shift left")
                if (commandToKeyMap.TryGetValue(parts[0], out ushort modifier1Key) &&
                    commandToKeyMap.TryGetValue(parts[1], out ushort modifier2Key) &&
                    commandToKeyMap.TryGetValue(parts[2], out ushort actionKey))
                {
                    // Send modifier keys down
                    SendKeyDown(modifier1Key);
                    SendKeyDown(modifier2Key);
                    
                    // Send the action key press
                    SendKeyDown(actionKey);
                    SendKeyUp(actionKey);
                    
                    // Send modifier keys up in reverse order
                    SendKeyUp(modifier2Key);
                    SendKeyUp(modifier1Key);
                }
            }
        }
        else if (commandToKeyMap.TryGetValue(recognizedText, out ushort keyCode))
        {
            // Send single key press
            SendKeyDown(keyCode);
            SendKeyUp(keyCode);
        }
    }
    
    private void SendKeyDown(ushort keyCode)
    {
        INPUT[] inputs = new INPUT[1];
        inputs[0].type = INPUT_KEYBOARD;
        inputs[0].union.keyboardInput.wVk = keyCode;
        inputs[0].union.keyboardInput.dwFlags = KEYEVENTF_KEYDOWN;
        inputs[0].union.keyboardInput.time = 0;
        inputs[0].union.keyboardInput.dwExtraInfo = IntPtr.Zero;
        
        SendInput(1, inputs, Marshal.SizeOf(typeof(INPUT)));
    }
    
    private void SendKeyUp(ushort keyCode)
    {
        INPUT[] inputs = new INPUT[1];
        inputs[0].type = INPUT_KEYBOARD;
        inputs[0].union.keyboardInput.wVk = keyCode;
        inputs[0].union.keyboardInput.dwFlags = KEYEVENTF_KEYUP;
        inputs[0].union.keyboardInput.time = 0;
        inputs[0].union.keyboardInput.dwExtraInfo = IntPtr.Zero;
        
        SendInput(1, inputs, Marshal.SizeOf(typeof(INPUT)));
    }

    private void ToggleRecognition(Button button)
    {
        if (isListening && recognizer != null)
        {
            // Stop recognition
            recognizer.RecognizeAsyncStop();
            isListening = false;
            button.Text = "Start Listening";
            
            if (this.Tag is Label statusLabel)
            {
                statusLabel.Text = "Status: Not Listening";
            }
        }
        else if (recognizer != null)
        {
            // Start recognition
            try
            {
                recognizer.RecognizeAsync(RecognizeMode.Multiple);
                isListening = true;
                button.Text = "Stop Listening";
                
                if (this.Tag is Label statusLabel)
                {
                    statusLabel.Text = "Status: Listening...";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting recognition: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
