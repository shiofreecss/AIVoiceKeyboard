using System;
using System.Collections.Generic;
using System.Speech.Recognition;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using System.Text.RegularExpressions;

namespace VoiceToKeyboard;

public partial class Form1 : Form
{
    private SpeechRecognitionEngine? commandRecognizer;
    private SpeechRecognitionEngine? dictationRecognizer;
    private bool isListening = false;
    private enum InputMode { Command, String }
    private InputMode currentMode = InputMode.Command;
    private StringBuilder typedText = new StringBuilder();
    private Label? modeIndicatorLabel;
    private TextBox? typedTextBox;
    private Label? recordingStatusLabel; // New recording status indicator
    private Label? speakingNowLabel; // Visual indicator for when to speak
    private Label? recognizedLabel; // Reference to the recognized text label
    private System.Windows.Forms.Timer? flashTimer; // Timer for flashing the speaking indicator
    
    // Whisper components
    private WhisperSpeechRecognition? whisperRecognition;
    private AudioCapture? audioCapture;
    private CancellationTokenSource? cancellationTokenSource;
    private bool useWhisperForString = true;
    private bool isWhisperReady = false;
    private CheckBox? useWhisperCheckbox;
    private bool isProcessingAudio = false; // Flag to prevent duplicate processing
    private string currentWhisperModel = "ggml-medium.bin"; // Default medium model
    private Button? resetModelButton; // Renamed from downloadModelButton
    
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
    const uint KEYEVENTF_UNICODE = 0x0004;

    // Dictionary to map voice commands to key codes
    private Dictionary<string, ushort> commandToKeyMap = new Dictionary<string, ushort>
    {
        // Mode switching commands
        { "command option", 0 },    // Special handling
        { "string option", 0 },     // Special handling
        { "command mode", 0 },      // Alias for command option
        { "string mode", 0 },       // Alias for string option
        
        // Basic navigation
        { "up", 0x26 },             // Up arrow
        { "down", 0x28 },           // Down arrow
        { "left", 0x25 },           // Left arrow
        { "right", 0x27 },          // Right arrow
        { "page up", 0x21 },        // Page Up
        { "page down", 0x22 },      // Page Down
        { "home", 0x24 },           // Home
        { "end", 0x23 },            // End
        { "top", 0x24 },            // Alias for Home
        { "bottom", 0x23 },         // Alias for End
        { "escape", 0x1B },         // Escape
        { "escape key", 0x1B },     // Alias for Escape
        { "esc", 0x1B },            // Alias for Escape
        { "enter", 0x0D },          // Enter
        { "return", 0x0D },         // Alias for Enter
        { "space", 0x20 },          // Space
        { "space bar", 0x20 },      // Alias for Space
        { "tab", 0x09 },            // Tab
        { "tab key", 0x09 },        // Alias for Tab
        { "delete", 0x2E },         // Delete
        { "delete key", 0x2E },     // Alias for Delete
        { "backspace", 0x08 },      // Backspace
        { "back space", 0x08 },     // Alias for Backspace

        // Function keys
        { "f1", 0x70 },             // F1
        { "f2", 0x71 },             // F2
        { "f3", 0x72 },             // F3
        { "f4", 0x73 },             // F4
        { "f5", 0x74 },             // F5
        { "f6", 0x75 },             // F6
        { "f7", 0x76 },             // F7
        { "f8", 0x77 },             // F8
        { "f9", 0x78 },             // F9
        { "f10", 0x79 },            // F10
        { "f11", 0x7A },            // F11
        { "f12", 0x7B },            // F12
        
        // Modifiers
        { "shift", 0x10 },          // Shift
        { "control", 0x11 },        // Control
        { "ctrl", 0x11 },           // Alias for Control
        { "alt", 0x12 },            // Alt
        { "windows", 0x5B },        // Windows key
        { "win", 0x5B },            // Alias for Windows
        
        // Common keyboard shortcuts
        { "copy", 0x43 },           // C (for Ctrl+C)
        { "paste", 0x56 },          // V (for Ctrl+V)
        { "cut", 0x58 },            // X (for Ctrl+X)
        { "undo", 0x5A },           // Z (for Ctrl+Z)
        { "redo", 0x59 },           // Y (for Ctrl+Y)
        { "select all", 0x41 },     // A (for Ctrl+A)
        { "save", 0x53 },           // S (for Ctrl+S)
        { "find", 0x46 },           // F (for Ctrl+F)
        { "print", 0x50 },          // P (for Ctrl+P)
        { "new", 0x4E },            // N (for Ctrl+N)
        { "open", 0x4F },           // O (for Ctrl+O)
        { "close", 0x57 },          // W (for Ctrl+W)
    };

    public Form1()
    {
        InitializeComponent();
        InitializeSpeechRecognition();
        // Don't initialize Whisper here, do it after the form is loaded
        InitializeUI();
        
        // Initialize the flash timer
        InitializeFlashTimer();
        
        // Add form load event handler to initialize Whisper after the window handle is created
        this.Load += Form1_Load;
    }

    private void Form1_Load(object sender, EventArgs e)
    {
        // Initialize Whisper after the form is fully loaded
        Task.Run(() => InitializeWhisper());
    }

    private void InitializeWhisper()
    {
        try
        {
            // Always use the medium model for stability
            currentWhisperModel = "ggml-medium.bin";
            
            UpdateStatusLabel("Initializing Whisper speech recognition...");
            
            // Create the Whisper recognition service with medium model only
            whisperRecognition = new WhisperSpeechRecognition("ggml-medium.bin");
            whisperRecognition.StatusChanged += (sender, status) => 
            {
                if (this.IsHandleCreated)
                {
                    UpdateStatusLabel(status);
                }
            };
            
            // Create the audio capture service
            audioCapture = new AudioCapture();
            audioCapture.StatusChanged += (sender, status) => 
            {
                if (this.IsHandleCreated)
                {
                    UpdateStatusLabel(status);
                }
            };
            audioCapture.AudioCaptured += async (sender, audioData) => 
            {
                if (whisperRecognition != null && isWhisperReady && currentMode == InputMode.String && useWhisperForString)
                {
                    await ProcessWhisperAudioAsync(audioData);
                }
            };
            
            // Initialize Whisper (this will download the model if needed)
            UpdateStatusLabel($"Loading medium model for best stability...");
            whisperRecognition.InitializeAsync().Wait();
            isWhisperReady = true;
            
            // Enable the checkbox when ready
            if (this.IsHandleCreated)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    if (useWhisperCheckbox != null)
                    {
                        useWhisperCheckbox.Enabled = isWhisperReady;
                    }
                });
            }
            
            UpdateStatusLabel("Whisper ready with medium model - start listening");
        }
        catch (AggregateException ex)
        {
            isWhisperReady = false;
            
            // Get inner exception if available (more detailed)
            Exception detailedException = ex.InnerException ?? ex;
            UpdateStatusLabel($"Whisper initialization error: {detailedException.Message}");
        }
        catch (Exception ex)
        {
            isWhisperReady = false;
            UpdateStatusLabel($"Whisper initialization error: {ex.Message}");
        }
    }

    private async Task ProcessWhisperAudioAsync(byte[] audioData)
    {
        if (whisperRecognition == null || !isWhisperReady)
            return;
            
        // Prevent duplicate processing
        if (isProcessingAudio)
        {
            UpdateStatusLabel("Still processing previous audio, skipping...");
            return;
        }
        
        isProcessingAudio = true;
        // No UpdateRecordingStatus here - we want to keep the indicator visible
        
        try
        {
            UpdateStatusLabel($"Processing word ({audioData.Length} bytes)...");
            
            // For real-time processing, we process even smaller audio segments
            // but we still want to filter out obvious silence
            if (audioData.Length < 500)
            {
                UpdateStatusLabel("Audio too short, ignoring");
                isProcessingAudio = false;
                // Keep recording status active - don't call UpdateRecordingStatus(false) here
                return;
            }
            
            // Set timeout for word processing to maintain responsiveness
            var recognitionTask = whisperRecognition.ProcessAudioAsync(audioData);
            string recognizedText = await await Task.WhenAny(
                recognitionTask, 
                Task.Delay(1500).ContinueWith(_ => string.Empty) // 1.5 second timeout
            );
            
            // If timeout occurred, handle gracefully
            if (string.IsNullOrEmpty(recognizedText) && !recognitionTask.IsCompleted)
            {
                UpdateStatusLabel("Word recognition timed out, continuing...");
                isProcessingAudio = false;
                return;
            }
            
            if (!string.IsNullOrEmpty(recognizedText))
            {
                // Handle the text processing on the UI thread
                this.Invoke((MethodInvoker)delegate {
                    HandleWhisperRecognizedText(recognizedText);
                });
            }
            else
            {
                // No text was recognized, prepare for next capture
                UpdateStatusLabel("No word detected, ready for next capture");
            }
        }
        catch (Exception ex)
        {
            UpdateStatusLabel($"Error processing audio: {ex.Message}");
            
            // Log the error details
            Console.WriteLine($"Whisper error: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
            
            if (ex.Message.Contains("not initialized") || ex.Message.Contains("Invalid"))
            {
                // Try to reinitialize Whisper with the same model (don't change the model)
                await RecoverWhisperProcessorAsync();
            }
        }
        finally
        {
            isProcessingAudio = false;
            // We don't change recording status here - let the StartListening loop manage it
        }
    }

    // Add a separate method to recover the processor without changing model
    private async Task RecoverWhisperProcessorAsync()
    {
        try
        {
            // Stop any ongoing processes
            isWhisperReady = false;
            
            if (whisperRecognition != null)
            {
                whisperRecognition.Dispose();
                whisperRecognition = null;
            }
            
            // Always use medium model for stability
            currentWhisperModel = "ggml-medium.bin";
            
            // Create new instance with medium model
            UpdateStatusLabel("Recovering Whisper processor with medium model...");
            whisperRecognition = new WhisperSpeechRecognition("ggml-medium.bin");
            whisperRecognition.StatusChanged += (sender, status) => 
            {
                if (this.IsHandleCreated)
                {
                    UpdateStatusLabel(status);
                }
            };
            
            // Initialize again with medium model
            await whisperRecognition.InitializeAsync();
            isWhisperReady = true;
            
            UpdateStatusLabel("Whisper processor recovered with medium model");
        }
        catch (Exception ex)
        {
            isWhisperReady = false;
            UpdateStatusLabel($"Whisper recovery error: {ex.Message}");
        }
    }

    private async Task ReinitializeWhisperAsync()
    {
        try
        {
            // Stop any ongoing processes
            isWhisperReady = false;
            
            if (whisperRecognition != null)
            {
                whisperRecognition.Dispose();
                whisperRecognition = null;
            }
            
            // Always use medium model for stability
            currentWhisperModel = "ggml-medium.bin";
            
            // Create new instance with medium model
            whisperRecognition = new WhisperSpeechRecognition("ggml-medium.bin");
            whisperRecognition.StatusChanged += (sender, status) => 
            {
                if (this.IsHandleCreated)
                {
                    UpdateStatusLabel(status);
                }
            };
            
            // Initialize again
            await whisperRecognition.InitializeAsync();
            isWhisperReady = true;
            
            UpdateStatusLabel("Initialized with medium model for best stability");
        }
        catch (Exception ex)
        {
            isWhisperReady = false;
            UpdateStatusLabel($"Whisper reinitialization error: {ex.Message}");
        }
    }

    private void ScrollToEnd(TextBox textBox)
    {
        // Select the end of the text to scroll there
        textBox.SelectionStart = textBox.Text.Length;
        textBox.SelectionLength = 0;
        textBox.ScrollToCaret();
    }

    private void HandleWhisperRecognizedText(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return;
        
        // Process the text - trim leading/trailing spaces and ensure proper capitalization
        text = text.Trim();
        
        // Filter out noise annotations like (keyboard tapping), [buzzer], <action>, etc.
        text = Regex.Replace(text, @"\([^)]*\)|\[[^\]]*\]|<[^>]*>", "").Trim();
        
        // Remove extra spaces that might have been created by removing annotations
        text = Regex.Replace(text, @"\s+", " ").Trim();
        
        // Skip if text is empty after filtering
        if (string.IsNullOrWhiteSpace(text))
            return;
        
        // Check for mode switching commands (case-insensitive)
        if (text.ToLower().Contains("command option") || text.ToLower().Equals("command mode"))
        {
            SwitchMode(InputMode.Command);
            return;
        }
        else if (text.ToLower().Contains("string option") || text.ToLower().Equals("string mode"))
        {
            SwitchMode(InputMode.String);
            return;
        }
        
        // Real-time word processing: Handle each word individually
        string[] words = text.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
        
        // First update the complete text in the UI
        typedText.Append(text + " ");
        if (typedTextBox != null)
        {
            typedTextBox.Text = typedText.ToString();
            ScrollToEnd(typedTextBox);
        }
        
        // Then process words individually for immediate feedback
        foreach (string word in words)
        {
            if (!string.IsNullOrWhiteSpace(word))
            {
                // Type the word character by character for immediate output
                foreach (char c in word)
                {
                    SendUnicodeChar(c);
                }
                
                // Add a space after each word
                SendUnicodeChar(' ');
                
                // Update status with each word for responsive feedback
                UpdateStatusLabel($"Word recognized: {word}");
            }
        }
    }

    private void UpdateStatusLabel(string status)
    {
        if (!this.IsHandleCreated)
            return;
        
        this.Invoke((MethodInvoker)delegate {
            if (this.Tag is Label statusLabel)
            {
                statusLabel.Text = status;
                
                // Highlight errors and warnings with different colors
                if (status.Contains("error", StringComparison.OrdinalIgnoreCase) || 
                    status.Contains("failed", StringComparison.OrdinalIgnoreCase) ||
                    status.Contains("not found", StringComparison.OrdinalIgnoreCase))
                {
                    statusLabel.ForeColor = System.Drawing.Color.Red;
                }
                else if (status.Contains("warning", StringComparison.OrdinalIgnoreCase) ||
                         status.Contains("download", StringComparison.OrdinalIgnoreCase))
                {
                    statusLabel.ForeColor = System.Drawing.Color.DarkOrange;
                }
                else
                {
                    statusLabel.ForeColor = System.Drawing.Color.FromArgb(80, 80, 80);
                }
                
                statusLabel.Refresh();
            }
        });
    }

    private void InitializeUI()
    {
        // Change form title and appearance
        this.Text = "AI Voice Keyboard - v1.0.0";
        this.BackColor = System.Drawing.Color.FromArgb(245, 247, 250);
        this.FormBorderStyle = FormBorderStyle.FixedSingle;
        this.MaximizeBox = false;
        this.Icon = new System.Drawing.Icon(System.Drawing.SystemIcons.Application, 40, 40);
        
        // Create panel for header
        Panel headerPanel = new Panel
        {
            BackColor = System.Drawing.Color.FromArgb(59, 130, 196),
            Dock = DockStyle.Top,
            Height = 60
        };
        
        // Add title label to header
        Label titleLabel = new Label
        {
            Text = "AI Voice Keyboard",
            ForeColor = System.Drawing.Color.White,
            Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold),
            AutoSize = true,
            Location = new System.Drawing.Point(20, 15)
        };
        headerPanel.Controls.Add(titleLabel);
        
        // Create main container panel
        Panel mainPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(20)
        };
        
        // Create left panel for controls
        Panel leftPanel = new Panel
        {
            Width = 220,
            Height = 350,
            Location = new System.Drawing.Point(20, 20),
            BackColor = System.Drawing.Color.White,
            BorderStyle = BorderStyle.None
        };
        
        // Add shadow effect
        leftPanel.Paint += (sender, e) => {
            ControlPaint.DrawBorder(e.Graphics, leftPanel.ClientRectangle,
                System.Drawing.Color.FromArgb(230, 230, 230), 1, ButtonBorderStyle.Solid,
                System.Drawing.Color.FromArgb(230, 230, 230), 1, ButtonBorderStyle.Solid,
                System.Drawing.Color.FromArgb(190, 190, 190), 2, ButtonBorderStyle.Solid,
                System.Drawing.Color.FromArgb(190, 190, 190), 2, ButtonBorderStyle.Solid);
        };
        
        // Create right panel for status and display
        Panel rightPanel = new Panel
        {
            Width = 360,
            Height = 350,
            Location = new System.Drawing.Point(250, 20),
            BackColor = System.Drawing.Color.White,
            BorderStyle = BorderStyle.None
        };
        
        // Add shadow effect
        rightPanel.Paint += (sender, e) => {
            ControlPaint.DrawBorder(e.Graphics, rightPanel.ClientRectangle,
                System.Drawing.Color.FromArgb(230, 230, 230), 1, ButtonBorderStyle.Solid,
                System.Drawing.Color.FromArgb(230, 230, 230), 1, ButtonBorderStyle.Solid,
                System.Drawing.Color.FromArgb(190, 190, 190), 2, ButtonBorderStyle.Solid,
                System.Drawing.Color.FromArgb(190, 190, 190), 2, ButtonBorderStyle.Solid);
        };
        
        // Mode selection group box
        GroupBox modeGroupBox = new GroupBox
        {
            Text = "Input Mode",
            Location = new System.Drawing.Point(15, 15),
            Size = new System.Drawing.Size(190, 80),
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            ForeColor = System.Drawing.Color.FromArgb(59, 130, 196)
        };
        
        // Create radio buttons for mode selection
        RadioButton commandModeRadio = new RadioButton
        {
            Text = "Command Mode",
            Checked = true,
            AutoSize = true,
            Location = new System.Drawing.Point(15, 25),
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            Cursor = Cursors.Hand
        };
        
        RadioButton stringModeRadio = new RadioButton
        {
            Text = "String Mode",
            AutoSize = true,
            Location = new System.Drawing.Point(15, 50),
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            Cursor = Cursors.Hand
        };
        
        // Add event handlers for radio buttons
        commandModeRadio.CheckedChanged += (sender, e) => 
        {
            if (commandModeRadio.Checked)
            {
                SwitchMode(InputMode.Command);
            }
        };
        
        stringModeRadio.CheckedChanged += (sender, e) => 
        {
            if (stringModeRadio.Checked)
            {
                SwitchMode(InputMode.String);
            }
        };
        
        modeGroupBox.Controls.Add(commandModeRadio);
        modeGroupBox.Controls.Add(stringModeRadio);
        leftPanel.Controls.Add(modeGroupBox);
        
        // Create a toggle button for starting/stopping recognition
        Button toggleButton = new Button
        {
            Text = "Start Listening",
            Size = new System.Drawing.Size(190, 50),
            Location = new System.Drawing.Point(15, 105),
            FlatStyle = FlatStyle.Flat,
            BackColor = System.Drawing.Color.FromArgb(59, 130, 196),
            ForeColor = System.Drawing.Color.White,
            Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        toggleButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(59, 130, 196);
        toggleButton.Click += (sender, e) => ToggleRecognition(toggleButton);
        leftPanel.Controls.Add(toggleButton);
        
        // Add a clear text button
        Button clearTextButton = new Button
        {
            Text = "Clear Text",
            Size = new System.Drawing.Size(190, 40),
            Location = new System.Drawing.Point(15, 165),
            FlatStyle = FlatStyle.Flat,
            BackColor = System.Drawing.Color.FromArgb(230, 230, 230),
            ForeColor = System.Drawing.Color.FromArgb(40, 40, 40),
            Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular),
            Cursor = Cursors.Hand
        };
        clearTextButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(200, 200, 200);
        
        clearTextButton.Click += (sender, e) => 
        {
            typedText.Clear();
            if (typedTextBox != null)
            {
                typedTextBox.Clear();
            }
        };
        leftPanel.Controls.Add(clearTextButton);
        
        // Add a copy text button
        Button copyTextButton = new Button
        {
            Text = "Copy Text",
            Size = new System.Drawing.Size(190, 40),
            Location = new System.Drawing.Point(15, 165 + 45), // Position below clear button
            FlatStyle = FlatStyle.Flat,
            BackColor = System.Drawing.Color.FromArgb(230, 230, 230),
            ForeColor = System.Drawing.Color.FromArgb(40, 40, 40),
            Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular),
            Cursor = Cursors.Hand
        };
        copyTextButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(200, 200, 200);
        
        copyTextButton.Click += (sender, e) => 
        {
            if (typedTextBox != null && !string.IsNullOrEmpty(typedTextBox.Text))
            {
                Clipboard.SetText(typedTextBox.Text);
                UpdateStatusLabel("Text copied to clipboard");
            }
        };
        leftPanel.Controls.Add(copyTextButton);
        
        // Add settings panel for speech recognition
        GroupBox settingsGroupBox = new GroupBox
        {
            Text = "Recognition Settings",
            Location = new System.Drawing.Point(15, 215 + 45), // Adjust position for new copy button
            Size = new System.Drawing.Size(190, 70),
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            ForeColor = System.Drawing.Color.FromArgb(59, 130, 196)
        };
        
        // Add checkbox for Whisper
        useWhisperCheckbox = new CheckBox
        {
            Text = "Use AI Recognition",
            Checked = useWhisperForString,
            Enabled = isWhisperReady,
            AutoSize = true,
            Location = new System.Drawing.Point(15, 25),
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            Cursor = Cursors.Hand
        };
        
        useWhisperCheckbox.CheckedChanged += (sender, e) => 
        {
            useWhisperForString = useWhisperCheckbox.Checked;
        };
        
        // Add Whisper model info as a simple label
        Label modelLabel = new Label
        {
            Text = "Whisper Model: Medium",
            AutoSize = true,
            Location = new System.Drawing.Point(15, 45),
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            ForeColor = System.Drawing.Color.FromArgb(80, 80, 80)
        };
        
        // Hidden reset model button (keeping the reference for code compatibility)
        resetModelButton = new Button
        {
            Text = "Reset Model",
            Size = new System.Drawing.Size(0, 0),
            Visible = false,
            Enabled = false
        };
        
        // Add components to settings group
        settingsGroupBox.Controls.Add(useWhisperCheckbox);
        settingsGroupBox.Controls.Add(modelLabel);
        settingsGroupBox.Controls.Add(resetModelButton);
        
        leftPanel.Controls.Add(settingsGroupBox);
        
        // Status panel
        Panel statusPanel = new Panel
        {
            Width = 340,
            Height = 70,
            Location = new System.Drawing.Point(10, 10),
            BackColor = System.Drawing.Color.FromArgb(245, 248, 250),
            BorderStyle = BorderStyle.FixedSingle
        };
        
        // Create a status label
        Label statusLabel = new Label
        {
            Text = "Status: Not Listening",
            AutoSize = false,
            Width = 320,
            Height = 25,
            Location = new System.Drawing.Point(10, 10),
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            ForeColor = System.Drawing.Color.FromArgb(80, 80, 80)
        };
        this.Tag = statusLabel; // Store reference to update later
        statusPanel.Controls.Add(statusLabel);
        
        // Create recording status indicator
        recordingStatusLabel = new Label
        {
            Text = "●",
            AutoSize = true,
            Location = new System.Drawing.Point(310, 10),
            Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Bold),
            ForeColor = System.Drawing.Color.Gray
        };
        statusPanel.Controls.Add(recordingStatusLabel);
        
        // Create mode indicator label
        modeIndicatorLabel = new Label
        {
            Text = "Mode: Command",
            AutoSize = true,
            Location = new System.Drawing.Point(10, 40),
            Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold),
            ForeColor = System.Drawing.Color.FromArgb(59, 130, 196)
        };
        statusPanel.Controls.Add(modeIndicatorLabel);
        rightPanel.Controls.Add(statusPanel);
        
        // Create recognized text label
        recognizedLabel = new Label
        {
            Text = "Recognized Text:",
            AutoSize = true,
            Location = new System.Drawing.Point(10, 90),
            Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular),
            ForeColor = System.Drawing.Color.FromArgb(80, 80, 80)
        };
        rightPanel.Controls.Add(recognizedLabel);
        
        // Create speaking indicator (initially hidden)
        speakingNowLabel = new Label
        {
            Text = "SPEAKING NOW - PLEASE TALK",
            AutoSize = false,
            Width = 340,
            Height = 30,
            Location = new System.Drawing.Point(10, 85),
            Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Bold),
            ForeColor = System.Drawing.Color.White,
            BackColor = System.Drawing.Color.FromArgb(0, 160, 0), // Green background
            TextAlign = ContentAlignment.MiddleCenter,
            Visible = false
        };
        rightPanel.Controls.Add(speakingNowLabel);
        
        // Create typed text display - make it larger
        typedTextBox = new TextBox
        {
            Text = "",
            AutoSize = false,
            Width = 340,
            Height = 220,
            Location = new System.Drawing.Point(10, 115),
            Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular),
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = System.Drawing.Color.White,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            WordWrap = true,
            Padding = new Padding(8)
        };
        
        // Add right-click context menu for copy operation
        ContextMenuStrip textContextMenu = new ContextMenuStrip();
        ToolStripMenuItem copyMenuItem = new ToolStripMenuItem("Copy");
        copyMenuItem.Click += (sender, e) => {
            if (!string.IsNullOrEmpty(typedTextBox.Text))
            {
                Clipboard.SetText(typedTextBox.Text);
                UpdateStatusLabel("Text copied to clipboard");
            }
        };
        textContextMenu.Items.Add(copyMenuItem);
        
        // Add select all menu item
        ToolStripMenuItem selectAllMenuItem = new ToolStripMenuItem("Select All");
        selectAllMenuItem.Click += (sender, e) => {
            typedTextBox.SelectAll();
        };
        textContextMenu.Items.Add(selectAllMenuItem);
        
        // Add clear text menu item
        ToolStripMenuItem clearMenuItem = new ToolStripMenuItem("Clear");
        clearMenuItem.Click += (sender, e) => {
            typedText.Clear();
            typedTextBox.Clear();
        };
        textContextMenu.Items.Add(clearMenuItem);
        
        typedTextBox.ContextMenuStrip = textContextMenu;
        
        rightPanel.Controls.Add(typedTextBox);
        
        // Instructions panel
        Panel instructionsPanel = new Panel
        {
            Width = 600,
            Height = 120,
            Location = new System.Drawing.Point(20, 380),
            BackColor = System.Drawing.Color.FromArgb(240, 248, 255),
            BorderStyle = BorderStyle.FixedSingle
        };
        
        // Create instructions header
        Label instructionsHeader = new Label
        {
            Text = "Instructions",
            AutoSize = true,
            Location = new System.Drawing.Point(10, 8),
            Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold),
            ForeColor = System.Drawing.Color.FromArgb(59, 130, 196)
        };
        instructionsPanel.Controls.Add(instructionsHeader);
        
        // Create instructions label
        Label instructionsLabel = new Label
        {
            Text = "• Command Mode: Say commands like \"up\", \"down\", \"enter\", \"control c\" for copy.\n" +
                  "• String Mode: Speak clearly for real-time word-by-word recognition.\n" +
                  "• Say \"command option\" or \"string option\" to switch between modes at any time.\n" +
                  "• Recording indicator: ● Gray = Not recording, ● Green = Recording in progress.\n" +
                  "• You can scroll through recognized text and copy it using the \"Copy Text\" button or right-click menu.",
            AutoSize = true,
            Location = new System.Drawing.Point(10, 30),
            Size = new System.Drawing.Size(580, 0),
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            ForeColor = System.Drawing.Color.FromArgb(60, 60, 60)
        };
        instructionsPanel.Controls.Add(instructionsLabel);
        
        // Add panels to main panel
        mainPanel.Controls.Add(leftPanel);
        mainPanel.Controls.Add(rightPanel);
        mainPanel.Controls.Add(instructionsPanel);
        
        // Add panels to form
        this.Controls.Add(mainPanel);
        this.Controls.Add(headerPanel);
        
        // Set form size
        this.ClientSize = new System.Drawing.Size(640, 580);
        this.StartPosition = FormStartPosition.CenterScreen;
    }

    private void InitializeSpeechRecognition()
    {
        try
        {
            // Create command recognition engine
            commandRecognizer = new SpeechRecognitionEngine();
            
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
            
            commandRecognizer.LoadGrammar(grammar);
            commandRecognizer.SpeechRecognized += CommandRecognizer_SpeechRecognized;
            
            // Set input to default audio device
            commandRecognizer.SetInputToDefaultAudioDevice();
            
            // Create dictation recognizer for string mode with enhanced settings
            dictationRecognizer = new SpeechRecognitionEngine();
            
            // Configure the dictation recognizer for better accuracy
            dictationRecognizer.InitialSilenceTimeout = TimeSpan.FromSeconds(2);
            dictationRecognizer.BabbleTimeout = TimeSpan.FromSeconds(1.5);
            dictationRecognizer.EndSilenceTimeout = TimeSpan.FromSeconds(0.5);
            dictationRecognizer.EndSilenceTimeoutAmbiguous = TimeSpan.FromSeconds(0.7);
            
            // Add dictation grammar with clear topic hint
            DictationGrammar dictationGrammar = new DictationGrammar();
            dictationGrammar.Name = "Dictation Grammar";
            dictationGrammar.Enabled = true;
            dictationRecognizer.LoadGrammar(dictationGrammar);
            
            // Also add the mode switching commands to dictation mode for reliability
            GrammarBuilder modeCommands = new GrammarBuilder();
            Choices modeChoices = new Choices(new string[] { "command option", "string option" });
            modeCommands.Append(modeChoices);
            Grammar modeGrammar = new Grammar(modeCommands);
            modeGrammar.Name = "Mode Commands";
            dictationRecognizer.LoadGrammar(modeGrammar);
            
            dictationRecognizer.SpeechRecognized += DictationRecognizer_SpeechRecognized;
            dictationRecognizer.SetInputToDefaultAudioDevice();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Speech recognition initialization error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void CommandRecognizer_SpeechRecognized(object? sender, SpeechRecognizedEventArgs e)
    {
        string recognizedText = e.Result.Text.ToLower();
        
        // Update status with recognized command
        if (this.Tag is Label statusLabel)
        {
            statusLabel.Text = $"Recognized: {recognizedText}";
            statusLabel.Refresh();
        }
        
        // Check for mode switching commands
        if (recognizedText == "string option" || recognizedText == "string mode")
        {
            SwitchMode(InputMode.String);
            return;
        }
        else if (recognizedText == "command option" || recognizedText == "command mode")
        {
            SwitchMode(InputMode.Command);
            return;
        }
        
        // Handle combined commands
        if (recognizedText.Contains(" "))
        {
            string[] parts = recognizedText.Split(' ');
            if (parts.Length == 2)
            {
                // Special case for Windows Speech Recognition sometimes misrecognizing modifiers
                if (parts[0] == "control" || parts[0] == "ctrl")
                {
                    // Check for common aliases like "control sea" instead of "control c"
                    string secondPart = parts[1];
                    if (secondPart == "sea" || secondPart == "see") secondPart = "c";
                    if (secondPart == "the") secondPart = "v";
                    if (secondPart == "eggs" || secondPart == "ex") secondPart = "x";
                    if (secondPart == "said") secondPart = "s";
                    if (secondPart == "eh") secondPart = "a";
                    
                    // Update parts array with corrected second part
                    parts[1] = secondPart;
                }
            
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
    
    private void DictationRecognizer_SpeechRecognized(object? sender, SpeechRecognizedEventArgs e)
    {
        // Only process if not using Whisper
        if (useWhisperForString && isWhisperReady)
            return;
            
        string recognizedText = e.Result.Text;
        
        // Filter out noise annotations like (keyboard tapping), [buzzer], <action>, etc.
        recognizedText = Regex.Replace(recognizedText, @"\([^)]*\)|\[[^\]]*\]|<[^>]*>", "").Trim();
        
        // Remove extra spaces that might have been created by removing annotations
        recognizedText = Regex.Replace(recognizedText, @"\s+", " ").Trim();
        
        // Skip if text is empty after filtering
        if (string.IsNullOrWhiteSpace(recognizedText))
            return;
        
        // Check for mode switching commands (allow these in either mode)
        if (recognizedText.ToLower() == "command option")
        {
            SwitchMode(InputMode.Command);
            return;
        }
        else if (recognizedText.ToLower() == "string option")
        {
            SwitchMode(InputMode.String);
            return;
        }
        
        // Update status with recognized text
        if (this.Tag is Label statusLabel)
        {
            statusLabel.Text = $"Recognized: {recognizedText}";
            statusLabel.Refresh();
        }
        
        // Split into words for real-time processing
        string[] words = recognizedText.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
        
        // First, update the complete text in the display
        typedText.Append(recognizedText + " ");
        if (typedTextBox != null)
        {
            typedTextBox.Text = typedText.ToString();
            ScrollToEnd(typedTextBox);
        }
        
        // Then process each word individually
        foreach (string word in words)
        {
            if (!string.IsNullOrWhiteSpace(word))
            {
                // Type the word character by character for real-time feedback
                foreach (char c in word)
                {
                    SendUnicodeChar(c);
                }
                
                // Add a space after each word
                SendUnicodeChar(' ');
                
                // Update status for each word
                UpdateStatusLabel($"Word recognized: {word}");
            }
        }
    }
    
    private void SendUnicodeChar(char c)
    {
        // Send a unicode character as keyboard input
        INPUT[] inputs = new INPUT[1];
        inputs[0].type = INPUT_KEYBOARD;
        inputs[0].union.keyboardInput.wVk = 0;
        inputs[0].union.keyboardInput.wScan = c;
        inputs[0].union.keyboardInput.dwFlags = KEYEVENTF_UNICODE;
        inputs[0].union.keyboardInput.time = 0;
        inputs[0].union.keyboardInput.dwExtraInfo = IntPtr.Zero;
        
        SendInput(1, inputs, Marshal.SizeOf(typeof(INPUT)));
        
        // Send key up
        inputs[0].union.keyboardInput.dwFlags = KEYEVENTF_KEYUP | KEYEVENTF_UNICODE;
        SendInput(1, inputs, Marshal.SizeOf(typeof(INPUT)));
        
        // Small delay to prevent characters from getting lost
        System.Threading.Thread.Sleep(10);
    }
    
    private void SwitchMode(InputMode mode)
    {
        StopListening();
        
        currentMode = mode;
        
        if (modeIndicatorLabel != null)
        {
            modeIndicatorLabel.Text = $"Mode: {mode}";
            modeIndicatorLabel.Refresh();
        }
        
        // Update radio buttons to reflect the current mode
        foreach (Control control in this.Controls)
        {
            if (control is Panel mainPanel)
            {
                foreach (Control panelControl in mainPanel.Controls)
                {
                    if (panelControl is Panel leftPanel)
                    {
                        foreach (Control leftPanelControl in leftPanel.Controls)
                        {
                            if (leftPanelControl is GroupBox modeGroupBox)
                            {
                                foreach (Control modeControl in modeGroupBox.Controls)
                                {
                                    if (modeControl is RadioButton radioButton)
                                    {
                                        if ((radioButton.Text == "Command Mode" && mode == InputMode.Command) ||
                                            (radioButton.Text == "String Mode" && mode == InputMode.String))
                                        {
                                            radioButton.Checked = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        
        if (isListening)
        {
            StartListening();
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

    private void StartListening()
    {
        try
        {
            // Start the appropriate recognition engine based on mode
            if (currentMode == InputMode.Command && commandRecognizer != null)
            {
                commandRecognizer.RecognizeAsync(RecognizeMode.Multiple);
                UpdateRecordingStatus(false); // Command mode doesn't show continuous recording
            }
            else if (currentMode == InputMode.String)
            {
                if (useWhisperForString && isWhisperReady && audioCapture != null)
                {
                    // Start Whisper-based recognition
                    cancellationTokenSource = new CancellationTokenSource();
                    Task.Run(async () => 
                    {
                        try 
                        {
                            UpdateStatusLabel("Listening for real-time words...");
                            
                            // More frequent recording cycles for real-time word processing
                            while (!cancellationTokenSource.Token.IsCancellationRequested)
                            {
                                // Show recording is active
                                UpdateRecordingStatus(true); 
                                
                                // Start recording session - this will stop automatically after word detection
                                await audioCapture.StartRecordingAsync(cancellationTokenSource.Token);
                                
                                // Only hide the indicator if we're done listening entirely
                                if (cancellationTokenSource.Token.IsCancellationRequested)
                                {
                                    UpdateRecordingStatus(false);
                                }
                                
                                // Smaller delay between recordings for more real-time feel
                                await Task.Delay(100, cancellationTokenSource.Token);
                            }
                        }
                        catch (TaskCanceledException)
                        {
                            // This is expected on cancellation
                            UpdateRecordingStatus(false);
                        }
                    }, cancellationTokenSource.Token);
                }
                else if (dictationRecognizer != null)
                {
                    // Fall back to Windows speech recognition
                    dictationRecognizer.RecognizeAsync(RecognizeMode.Multiple);
                    UpdateRecordingStatus(true); // Show recording active
                }
            }
            
            isListening = true;
            UpdateStatusLabel($"Listening in {currentMode} mode...");
        }
        catch (Exception ex)
        {
            UpdateStatusLabel($"Error starting recognition: {ex.Message}");
            MessageBox.Show($"Error starting recognition: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
    
    private void StopListening()
    {
        // Stop command recognizer if active
        if (commandRecognizer != null && currentMode == InputMode.Command)
        {
            commandRecognizer.RecognizeAsyncStop();
        }
        
        // Stop string mode recognition
        if (currentMode == InputMode.String)
        {
            if (useWhisperForString && audioCapture != null)
            {
                // Stop Whisper-based recognition
                cancellationTokenSource?.Cancel();
                audioCapture.StopRecordingAsync().Wait();
            }
            else if (dictationRecognizer != null)
            {
                // Stop Windows speech recognition
                dictationRecognizer.RecognizeAsyncStop();
            }
        }
        
        // Update recording status to off
        UpdateRecordingStatus(false);
    }

    private void ToggleRecognition(Button button)
    {
        if (isListening)
        {
            // Stop recognition
            StopListening();
            isListening = false;
            button.Text = "Start Listening";
            button.BackColor = System.Drawing.Color.FromArgb(59, 130, 196);
            
            // Update UI
            UpdateStatusLabel("Status: Not Listening");
            UpdateRecordingStatus(false);
        }
        else
        {
            // Start recognition
            StartListening();
            button.Text = "Stop Listening";
            button.BackColor = System.Drawing.Color.FromArgb(192, 0, 0);
            
            // Update UI based on mode
            if (currentMode == InputMode.Command)
            {
                UpdateStatusLabel("Listening for commands...");
            }
            else
            {
                UpdateStatusLabel("Listening for speech...");
                if (useWhisperForString && isWhisperReady)
                {
                    UpdateStatusLabel("Listening using Whisper (speak clearly)...");
                }
                else
                {
                    UpdateStatusLabel("Listening using Windows Speech Recognition...");
                }
            }
        }
    }

    private void UpdateRecordingStatus(bool isRecording)
    {
        if (!this.IsHandleCreated)
            return;
            
        this.Invoke((MethodInvoker)delegate {
            // Update recording status dot
            if (recordingStatusLabel != null)
            {
                // Update color based on recording status
                recordingStatusLabel.ForeColor = isRecording 
                    ? System.Drawing.Color.FromArgb(0, 160, 0) // Green dot for recording
                    : System.Drawing.Color.Gray;
                
                // Add a tooltip to explain the indicator
                ToolTip tooltip = new ToolTip();
                tooltip.SetToolTip(recordingStatusLabel, isRecording 
                    ? "Recording in progress..." 
                    : "Not recording");
                
                recordingStatusLabel.Refresh();
            }
            
            // Update speaking now indicator
            if (speakingNowLabel != null && currentMode == InputMode.String)
            {
                // Only show the speaking indicator in String mode
                speakingNowLabel.Visible = isRecording;
                
                // Start or stop the flash timer
                if (isRecording && flashTimer != null)
                {
                    flashTimer.Start();
                    
                    // Set initial colors
                    speakingNowLabel.BackColor = System.Drawing.Color.FromArgb(0, 160, 0);
                    speakingNowLabel.Text = "● SPEAKING NOW - PLEASE TALK ●";
                }
                else if (flashTimer != null)
                {
                    flashTimer.Stop();
                }
                
                // Ensure the text label is updated to avoid overlap
                if (recognizedLabel != null)
                {
                    recognizedLabel.Visible = !isRecording;
                }
            }
            else if (speakingNowLabel != null)
            {
                // Always hide in Command mode
                speakingNowLabel.Visible = false;
                
                // Stop the flash timer
                if (flashTimer != null)
                {
                    flashTimer.Stop();
                }
            }
        });
    }

    private void InitializeFlashTimer()
    {
        // Create and configure the timer for flashing the speaking indicator
        flashTimer = new System.Windows.Forms.Timer();
        flashTimer.Interval = 500; // Flash every half second
        flashTimer.Tick += (sender, e) => {
            if (speakingNowLabel != null && speakingNowLabel.Visible)
            {
                // Toggle between green and a darker green for the flash effect
                if (speakingNowLabel.BackColor == System.Drawing.Color.FromArgb(0, 160, 0))
                {
                    speakingNowLabel.BackColor = System.Drawing.Color.FromArgb(0, 120, 0);
                    speakingNowLabel.Text = "SPEAKING NOW - PLEASE TALK";
                }
                else
                {
                    speakingNowLabel.BackColor = System.Drawing.Color.FromArgb(0, 160, 0);
                    speakingNowLabel.Text = "● SPEAKING NOW - PLEASE TALK ●";
                }
            }
        };
    }

    // Clean up resources
    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        base.OnFormClosing(e);
        
        // Clean up resources
        whisperRecognition?.Dispose();
        audioCapture?.Dispose();
        cancellationTokenSource?.Dispose();
        
        // Clean up the flash timer
        if (flashTimer != null)
        {
            flashTimer.Stop();
            flashTimer.Dispose();
            flashTimer = null;
        }
    }
}
