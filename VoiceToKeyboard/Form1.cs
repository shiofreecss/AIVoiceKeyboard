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
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace VoiceToKeyboard;

// Floating overlay form that stays on top of other windows
public class OverlayForm : Form
{
    private Button recordButton;
    private Button closeButton;
    private Form1 parentForm;
    private bool isDragging = false;
    private Point dragStartPoint;
    private NotifyIcon? notifyIcon;
    private Label modeIndicator;
    private Panel dragHandle; // Add a drag handle panel

    // P/Invoke for setting window to be click-through when needed
    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
    
    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);
    
    private const int GWL_EXSTYLE = -20;
    private const int WS_EX_LAYERED = 0x80000;
    private const int WS_EX_TRANSPARENT = 0x20;
    
    public OverlayForm(Form1 parent)
    {
        parentForm = parent;
        
        // Basic form settings
        this.FormBorderStyle = FormBorderStyle.None;
        this.ShowInTaskbar = false;
        this.TopMost = true;
        this.BackColor = Color.FromArgb(31, 34, 52); // Match app theme
        this.Size = new Size(60, 100); // Increase height to accommodate drag handle
        this.StartPosition = FormStartPosition.Manual;
        this.Opacity = 0.9; // Less transparent for better visibility
        
        // Position it at the right side of the screen
        Rectangle screen = Screen.PrimaryScreen.WorkingArea;
        this.Location = new Point(screen.Width - this.Width - 20, screen.Height / 2 - this.Height / 2);
        
        // Create the drag handle
        dragHandle = new Panel();
        dragHandle.Size = new Size(50, 8);
        dragHandle.Location = new Point(5, 0);
        dragHandle.BackColor = Color.FromArgb(51, 54, 72); // More subtle dark gray instead of blue
        dragHandle.Cursor = Cursors.SizeAll;
        this.Controls.Add(dragHandle);
        
        // Create the record button (adjust location for drag handle)
        recordButton = new Button();
        recordButton.Size = new Size(50, 50);
        recordButton.Location = new Point(5, 10); // Move down to accommodate drag handle
        recordButton.FlatStyle = FlatStyle.Flat;
        recordButton.FlatAppearance.BorderSize = 0;
        recordButton.BackColor = Color.Gray; // Default to gray (not recording)
        recordButton.ForeColor = Color.White;
        recordButton.Text = "●";
        recordButton.Font = new Font("Arial", 16, FontStyle.Bold);
        recordButton.Click += RecordButton_Click;
        this.Controls.Add(recordButton);
        
        // Create mode indicator label (adjust location)
        modeIndicator = new Label();
        modeIndicator.Size = new Size(50, 25);
        modeIndicator.Location = new Point(5, 65); // Adjust for drag handle
        modeIndicator.BackColor = Color.FromArgb(41, 44, 62);
        modeIndicator.ForeColor = Color.White;
        modeIndicator.Text = "CMD";
        modeIndicator.TextAlign = ContentAlignment.MiddleCenter;
        modeIndicator.Font = new Font("Arial", 8, FontStyle.Bold);
        modeIndicator.Cursor = Cursors.Hand;
        modeIndicator.Click += ModeIndicator_Click;
        this.Controls.Add(modeIndicator);
        
        // Create a small close button (X) in top-right corner
        closeButton = new Button();
        closeButton.Size = new Size(18, 18);
        closeButton.Location = new Point(this.Width - 21, 0); // Reposition to top-right corner
        closeButton.FlatStyle = FlatStyle.Flat;
        closeButton.FlatAppearance.BorderSize = 0;
        closeButton.BackColor = Color.FromArgb(192, 0, 0); // Red for close button
        closeButton.ForeColor = Color.White;
        closeButton.Text = "×";
        closeButton.Font = new Font("Arial", 10, FontStyle.Bold);
        closeButton.TextAlign = ContentAlignment.MiddleCenter;
        closeButton.Cursor = Cursors.Hand;
        closeButton.Click += CloseButton_Click;
        this.Controls.Add(closeButton);
        
        // Make the form shape rounded rectangular with smoother corners
        GraphicsPath path = new GraphicsPath();
        int radius = 15; // Corner radius
        
        // Add rounded rectangle for main form
        path.AddArc(0, 0, radius * 2, radius * 2, 180, 90); // Top-left corner
        path.AddArc(this.Width - radius * 2, 0, radius * 2, radius * 2, 270, 90); // Top-right corner
        path.AddArc(this.Width - radius * 2, this.Height - radius * 2, radius * 2, radius * 2, 0, 90); // Bottom-right corner
        path.AddArc(0, this.Height - radius * 2, radius * 2, radius * 2, 90, 90); // Bottom-left corner
        path.CloseAllFigures();
        
        this.Region = new Region(path);
        
        // Bring controls to front to ensure visibility with rounded corners
        dragHandle.BringToFront();
        closeButton.BringToFront();
        
        // Add drag handle specific mouse events
        dragHandle.MouseDown += DragHandle_MouseDown;
        dragHandle.MouseMove += DragHandle_MouseMove;
        dragHandle.MouseUp += DragHandle_MouseUp;
        
        // Modify tooltip for dragHandle to make it clear that this is for dragging
        ToolTip dragTooltip = new ToolTip();
        dragTooltip.SetToolTip(dragHandle, "Drag to move");
        
        // Mouse events for dragging (keep existing events)
        this.MouseDown += OverlayForm_MouseDown;
        this.MouseMove += OverlayForm_MouseMove;
        this.MouseUp += OverlayForm_MouseUp;
        recordButton.MouseDown += OverlayForm_MouseDown;
        recordButton.MouseMove += OverlayForm_MouseMove;
        recordButton.MouseUp += OverlayForm_MouseUp;
        
        // Create context menu for right-click on the button
        ContextMenuStrip contextMenu = new ContextMenuStrip();
        
        // Add menu items
        ToolStripMenuItem commandModeItem = new ToolStripMenuItem("Command Mode");
        commandModeItem.Click += (s, e) => parentForm.SwitchModeFromOverlay(Form1.InputMode.Command);
        contextMenu.Items.Add(commandModeItem);
        
        ToolStripMenuItem stringModeItem = new ToolStripMenuItem("String Mode");
        stringModeItem.Click += (s, e) => parentForm.SwitchModeFromOverlay(Form1.InputMode.String);
        contextMenu.Items.Add(stringModeItem);
        
        contextMenu.Items.Add(new ToolStripSeparator());
        
        ToolStripMenuItem openMainItem = new ToolStripMenuItem("Open Main Window");
        openMainItem.Click += (s, e) => parentForm.ShowMainWindow();
        contextMenu.Items.Add(openMainItem);
        
        ToolStripMenuItem exitItem = new ToolStripMenuItem("Exit Application");
        exitItem.Click += (s, e) => parentForm.ExitApplication();
        contextMenu.Items.Add(exitItem);
        
        this.ContextMenuStrip = contextMenu;
        recordButton.ContextMenuStrip = contextMenu;
        modeIndicator.ContextMenuStrip = contextMenu;
        
        // Add notifyIcon for tray access
        notifyIcon = new NotifyIcon();
        notifyIcon.Text = "AI Voice Keyboard";
        try
        {
            string iconPath = Path.Combine(Application.StartupPath, "icon", "ai-voice-keyboard.ico");
            if (File.Exists(iconPath))
            {
                notifyIcon.Icon = new Icon(iconPath);
            }
            else
            {
                iconPath = Path.Combine("icon", "ai-voice-keyboard.ico");
                if (File.Exists(iconPath))
                {
                    notifyIcon.Icon = new Icon(iconPath);
                }
                else
                {
                    notifyIcon.Icon = SystemIcons.Application;
                }
            }
        }
        catch
        {
            notifyIcon.Icon = SystemIcons.Application;
        }
        notifyIcon.Visible = true;
        notifyIcon.ContextMenuStrip = contextMenu;
        notifyIcon.DoubleClick += (s, e) => parentForm.ShowMainWindow();
    }
    
    // Handler for the close button
    private void CloseButton_Click(object sender, EventArgs e)
    {
        // Hide overlay instead of exiting the application
        parentForm.HideOverlayAndUntickCheckbox();
    }
    
    // Override CreateParams to ensure the overlay stays on top across desktops
    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            // Add the always-on-top style
            cp.ExStyle |= 0x00000008; // WS_EX_TOPMOST
            // Add the tool window style to reduce taskbar presence
            cp.ExStyle |= 0x00000080; // WS_EX_TOOLWINDOW
            return cp;
        }
    }
    
    // Mode indicator click event handler
    private void ModeIndicator_Click(object sender, EventArgs e)
    {
        // Toggle between Command and String modes
        if (modeIndicator.Text == "CMD")
        {
            parentForm.SwitchModeFromOverlay(Form1.InputMode.String);
        }
        else
        {
            parentForm.SwitchModeFromOverlay(Form1.InputMode.Command);
        }
    }
    
    // Update button appearance based on recording state and current mode
    public void UpdateButtonState(bool isRecording, Form1.InputMode currentMode)
    {
        // Use BeginInvoke to ensure we're on the UI thread
        if (this.IsHandleCreated && !this.IsDisposed)
        {
            this.BeginInvoke((MethodInvoker)delegate
            {
                // Update recording button
                if (isRecording)
                {
                    recordButton.BackColor = Color.Red;
                    recordButton.Text = "■";
                    
                    if (notifyIcon != null)
                        notifyIcon.Text = $"AI Voice Keyboard ({currentMode} Mode - Recording)";
                }
                else
                {
                    recordButton.BackColor = Color.Gray;
                    recordButton.Text = "●";
                    
                    if (notifyIcon != null)
                        notifyIcon.Text = $"AI Voice Keyboard ({currentMode} Mode - Not Recording)";
                }
                
                // Update mode indicator
                if (currentMode == Form1.InputMode.Command)
                {
                    modeIndicator.Text = "CMD";
                    modeIndicator.BackColor = Color.FromArgb(41, 44, 62);
                }
                else
                {
                    modeIndicator.Text = "STR";
                    modeIndicator.BackColor = Color.FromArgb(84, 130, 210); // Bitwarden blue
                }
            });
        }
    }
    
    private void RecordButton_Click(object sender, EventArgs e)
    {
        // Toggle recording state in parent form
        parentForm.ToggleRecordingFromOverlay();
    }
    
    // Handle dragging of the overlay
    private void OverlayForm_MouseDown(object sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            isDragging = true;
            dragStartPoint = new Point(e.X, e.Y);
        }
    }
    
    private void OverlayForm_MouseMove(object sender, MouseEventArgs e)
    {
        if (isDragging)
        {
            Point currentPoint = PointToScreen(new Point(e.X, e.Y));
            this.Location = new Point(
                currentPoint.X - dragStartPoint.X,
                currentPoint.Y - dragStartPoint.Y
            );
        }
    }
    
    private void OverlayForm_MouseUp(object sender, MouseEventArgs e)
    {
        isDragging = false;
    }
    
    // Make form click-through when needed
    public void SetClickThrough(bool clickThrough)
    {
        if (this.IsHandleCreated && !this.IsDisposed)
        {
            try
            {
                if (clickThrough)
                {
                    int exStyle = GetWindowLong(this.Handle, GWL_EXSTYLE);
                    SetWindowLong(this.Handle, GWL_EXSTYLE, exStyle | WS_EX_LAYERED | WS_EX_TRANSPARENT);
                }
                else
                {
                    int exStyle = GetWindowLong(this.Handle, GWL_EXSTYLE);
                    SetWindowLong(this.Handle, GWL_EXSTYLE, exStyle & ~(WS_EX_LAYERED | WS_EX_TRANSPARENT));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error setting click-through: {ex.Message}");
            }
        }
    }
    
    // Clean up notifyIcon when the form is closing
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            if (notifyIcon != null)
            {
                notifyIcon.Dispose();
                notifyIcon = null;
            }
        }
        base.Dispose(disposing);
    }

    // Mouse event handlers specifically for the drag handle
    private void DragHandle_MouseDown(object sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            isDragging = true;
            
            // Convert drag handle coordinates to form coordinates
            Point controlPoint = dragHandle.PointToScreen(new Point(e.X, e.Y));
            dragStartPoint = this.PointToClient(controlPoint);
        }
    }
    
    private void DragHandle_MouseMove(object sender, MouseEventArgs e)
    {
        if (isDragging)
        {
            // Convert the current coordinates to screen coordinates
            Point controlPoint = dragHandle.PointToScreen(new Point(e.X, e.Y));
            Point formPoint = this.PointToClient(controlPoint);
            
            // Move the form
            this.Location = new Point(
                this.Location.X + (formPoint.X - dragStartPoint.X),
                this.Location.Y + (formPoint.Y - dragStartPoint.Y)
            );
        }
    }
    
    private void DragHandle_MouseUp(object sender, MouseEventArgs e)
    {
        isDragging = false;
    }
}

public partial class Form1 : Form
{
    private SpeechRecognitionEngine? commandRecognizer;
    private SpeechRecognitionEngine? dictationRecognizer;
    private bool isListening = false;
    public enum InputMode { Command, String }
    private InputMode currentMode = InputMode.Command;
    private StringBuilder typedText = new StringBuilder();
    private Label? modeIndicatorLabel;
    private TextBox? typedTextBox;
    private Label? recordingStatusLabel; // New recording status indicator
    private Label? speakingNowLabel; // Visual indicator for when to speak
    private Label? recognizedLabel; // Reference to the recognized text label
    private System.Windows.Forms.Timer? flashTimer; // Timer for flashing the speaking indicator
    
    // Version information
    public string appVersion = "1.0.2"; // Default version if config file not found
    public string buildNumber = "1002";
    public string releaseDate = "2025-03-30";
    
    // Whisper components
    private WhisperSpeechRecognition? whisperRecognition;
    private AudioCapture? audioCapture;
    private CancellationTokenSource? cancellationTokenSource;
    private bool useWhisperForString = true;
    private bool isWhisperReady = false;
    private Label? useWhisperLabel;
    private bool isProcessingAudio = false; // Flag to prevent duplicate processing
    private string currentWhisperModel = "ggml-medium.bin"; // Default medium model
    private Button? resetModelButton; // Renamed from downloadModelButton
    
    // Audio constants for processing time calculation
    private const int SampleRate = 16000;     // 16kHz is optimal for Whisper
    private const int BitsPerSample = 16;     // 16-bit audio
    private const int Channels = 1;           // Mono

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

    private OverlayForm? overlayForm;
    private bool overlayEnabled = false;
    private CheckBox? enableOverlayCheckBox;
    
    // Add flag to track if we've shown the main window
    private bool mainWindowShown = true;

    public Form1()
    {
        InitializeComponent();
        LoadVersionInfo();
        InitializeSpeechRecognition();
        // Don't initialize Whisper here, do it after the form is loaded
        InitializeUI();
        
        // Initialize the flash timer
        InitializeFlashTimer();
        
        // Add form load event handler to initialize Whisper after the window handle is created
        this.Load += Form1_Load;
        
        // Add event handler for form resize to handle minimizing
        this.Resize += Form1_Resize;
    }

    private void LoadVersionInfo()
    {
        try
        {
            // Try to load version info from config file
            string configPath = Path.Combine(Application.StartupPath, "version.cfg");
            
            // If not found in the app directory, try the project directory
            if (!File.Exists(configPath))
            {
                configPath = Path.Combine("version.cfg");
            }
            
            if (File.Exists(configPath))
            {
                string[] lines = File.ReadAllLines(configPath);
                foreach (string line in lines)
                {
                    // Skip empty lines and comments
                    if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                        continue;
                        
                    // Parse key-value pairs
                    string[] parts = line.Split('=');
                    if (parts.Length == 2)
                    {
                        string key = parts[0].Trim();
                        string value = parts[1].Trim();
                        
                        if (key.Equals("version", StringComparison.OrdinalIgnoreCase))
                            appVersion = value;
                        else if (key.Equals("build", StringComparison.OrdinalIgnoreCase))
                            buildNumber = value;
                        else if (key.Equals("release_date", StringComparison.OrdinalIgnoreCase))
                            releaseDate = value;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // If any error occurs, default values will be used
            Console.WriteLine($"Error loading version info: {ex.Message}");
        }
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
            
            // Update the label when ready
            if (this.IsHandleCreated)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    if (useWhisperLabel != null)
                    {
                        useWhisperLabel.Enabled = isWhisperReady;
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
            UpdateStatusLabel($"Processing speech ({audioData.Length / 1024} KB)...");
            
            // For real-time processing, but we still want to filter out obvious silence
            if (audioData.Length < 500)
            {
                UpdateStatusLabel("Audio too short, ignoring");
                isProcessingAudio = false;
                // Keep recording status active - don't call UpdateRecordingStatus(false) here
                return;
            }
            
            // Calculate an appropriate timeout based on audio length
            // Roughly 1 second processing time per 2 seconds of audio, with minimum 2 seconds
            int audioLengthMs = audioData.Length / (SampleRate * BitsPerSample * Channels / 8 / 1000);
            int timeoutMs = Math.Max(2000, audioLengthMs / 2);
            
            // Set timeout for speech processing to maintain responsiveness
            var recognitionTask = whisperRecognition.ProcessAudioAsync(audioData);
            string recognizedText = await await Task.WhenAny(
                recognitionTask, 
                Task.Delay(timeoutMs).ContinueWith(_ => string.Empty) // Dynamic timeout
            );
            
            // If timeout occurred, handle gracefully
            if (string.IsNullOrEmpty(recognizedText) && !recognitionTask.IsCompleted)
            {
                UpdateStatusLabel("Speech recognition timed out, continuing...");
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
                UpdateStatusLabel("No speech detected, ready for next capture");
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
        
        // Process the text - trim leading/trailing spaces
        text = text.Trim();
        
        // Filter out noise annotations like (keyboard tapping), [buzzer], <action>, etc.
        text = Regex.Replace(text, @"\([^)]*\)|\[[^\]]*\]|<[^>]*>", "").Trim();
        
        // Filter out asterisk-enclosed noise annotations like *cough*, *sad music*, etc.
        text = Regex.Replace(text, @"\*[^*]*\*", "").Trim();
        
        // Filter out common sound effect markers in various formats
        text = Regex.Replace(text, @"\{[^}]*\}", "").Trim(); // {sound effects}
        text = Regex.Replace(text, @"#[^#]*#", "").Trim();   // #sound effects#
        text = Regex.Replace(text, @"~[^~]*~", "").Trim();   // ~background noise~
        
        // Filter out common sound effect words with surrounding punctuation
        string[] noisePhrases = new[] { 
            "\\bcough\\b", "\\bsigh\\b", "\\blaugh\\b", "\\bgasp\\b", 
            "\\bmusic\\b", "\\bapplause\\b", "\\bsniff\\b", "\\bchuckle\\b", 
            "\\blaughter\\b", "\\bbackground noise\\b", "\\bnoise\\b", "\\bsound\\b",
            "\\bclears throat\\b", "\\bhumming\\b", "\\bexhales\\b", "\\binhales\\b"
        };
        
        foreach (string phrase in noisePhrases)
        {
            // Remove annotations with the phrases in various formats
            text = Regex.Replace(text, $"\\(.*{phrase}.*\\)", "", RegexOptions.IgnoreCase).Trim();
            text = Regex.Replace(text, $"\\[.*{phrase}.*\\]", "", RegexOptions.IgnoreCase).Trim();
            text = Regex.Replace(text, $"\\{{.*{phrase}.*\\}}", "", RegexOptions.IgnoreCase).Trim();
            text = Regex.Replace(text, $"\\*.*{phrase}.*\\*", "", RegexOptions.IgnoreCase).Trim();
            
            // Also try to remove the standalone terms with optional punctuation
            text = Regex.Replace(text, $"\\W*{phrase}\\W*", " ", RegexOptions.IgnoreCase).Trim();
        }
        
        // Remove extra spaces that might have been created by removing annotations
        text = Regex.Replace(text, @"\s+", " ").Trim();
        
        // Clean up and fix punctuation
        text = CleanupText(text);
        
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
        
        // Add the recognized text to the history
        typedText.Append(text + " ");
        if (typedTextBox != null)
        {
            typedTextBox.Text = typedText.ToString();
            ScrollToEnd(typedTextBox);
        }
        
        // Process the text for typing
        bool isFirstWord = true;
        string[] sentences = text.Split(new[] { '.', '!', '?' }, StringSplitOptions.RemoveEmptyEntries);
        
        foreach (string sentence in sentences)
        {
            if (string.IsNullOrWhiteSpace(sentence))
                continue;
                
            // Split into words
            string[] words = sentence.Trim().Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (string word in words)
            {
                if (string.IsNullOrWhiteSpace(word))
                    continue;
                
                // Add a space between words (but not before the first word)
                if (!isFirstWord)
                {
                    SendUnicodeChar(' ');
                }
                
                // Type the word character by character
                foreach (char c in word)
                {
                    SendUnicodeChar(c);
                }
                
                isFirstWord = false;
            }
            
            // Add sentence end punctuation if there are multiple sentences
            if (sentences.Length > 1)
            {
                SendUnicodeChar('.');
                SendUnicodeChar(' ');
            }
        }
        
        // Update status with recognized text summary
        if (text.Length > 30)
        {
            UpdateStatusLabel($"Speech recognized: \"{text.Substring(0, 27)}...\"");
        }
        else
        {
            UpdateStatusLabel($"Speech recognized: \"{text}\"");
        }
    }

    private string CleanupText(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;
            
        // Ensure proper spacing around punctuation
        text = Regex.Replace(text, @"\s+([.,!?:;])", "$1");
        
        // Ensure first letter is capitalized
        if (text.Length > 0)
        {
            text = char.ToUpper(text[0]) + text.Substring(1);
        }
        
        // Fix common formatting issues
        
        // Fix "i" -> "I" (standalone pronoun)
        text = Regex.Replace(text, @"\bi\b", "I");
        
        // Ensure proper spacing after punctuation
        text = Regex.Replace(text, @"([.,!?:;])(\S)", "$1 $2");
        
        // Remove leading/trailing quotes and parentheses if they're not properly closed
        text = BalanceDelimiters(text);
        
        return text;
    }
    
    private string BalanceDelimiters(string text)
    {
        // Check for mismatched quotes
        int singleQuotes = text.Count(c => c == '\'');
        int doubleQuotes = text.Count(c => c == '\"');
        
        // Remove trailing single quote if unbalanced
        if (singleQuotes % 2 != 0 && text.EndsWith("'"))
        {
            text = text.Substring(0, text.Length - 1);
        }
        
        // Remove trailing double quote if unbalanced
        if (doubleQuotes % 2 != 0 && text.EndsWith("\""))
        {
            text = text.Substring(0, text.Length - 1);
        }
        
        // Check for mismatched parentheses
        int openParens = text.Count(c => c == '(');
        int closeParens = text.Count(c => c == ')');
        
        // Remove trailing close parenthesis if unbalanced
        if (openParens < closeParens && text.EndsWith(")"))
        {
            text = text.Substring(0, text.Length - 1);
        }
        
        return text;
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
                    statusLabel.ForeColor = System.Drawing.Color.FromArgb(255, 100, 100); // Brighter red for errors
                }
                else if (status.Contains("warning", StringComparison.OrdinalIgnoreCase) ||
                         status.Contains("download", StringComparison.OrdinalIgnoreCase))
                {
                    statusLabel.ForeColor = System.Drawing.Color.FromArgb(255, 180, 80); // Brighter orange for warnings
                }
                else
                {
                    statusLabel.ForeColor = System.Drawing.Color.FromArgb(123, 162, 217); // Lighter blue for normal status
                }
                
                statusLabel.Refresh();
            }
        });
    }

    private void InitializeUI()
    {
        // Change form title and appearance
        this.Text = $"AI Voice Keyboard - v{appVersion}";
        this.BackColor = System.Drawing.Color.FromArgb(20, 21, 30); // Dark background like Bitwarden
        this.FormBorderStyle = FormBorderStyle.FixedSingle;
        this.MaximizeBox = false;
        
        // Try to load the custom icon, fall back to system icon if not found
        try
        {
            string iconPath = Path.Combine(Application.StartupPath, "icon", "ai-voice-keyboard.ico");
            if (File.Exists(iconPath))
            {
                this.Icon = new System.Drawing.Icon(iconPath);
            }
            else
            {
                // Try relative path from executable directory
                iconPath = Path.Combine("icon", "ai-voice-keyboard.ico");
                if (File.Exists(iconPath))
                {
                    this.Icon = new System.Drawing.Icon(iconPath);
                }
                else
                {
                    // Fallback to system icon
                    this.Icon = System.Drawing.SystemIcons.Application;
                }
            }
        }
        catch
        {
            // Fallback to system icon on any error
            this.Icon = System.Drawing.SystemIcons.Application;
        }
        
        // Create panel for header
        Panel headerPanel = new Panel
        {
            BackColor = System.Drawing.Color.FromArgb(31, 34, 52), // Dark header like Bitwarden
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
        
        // Add developer link to bottom right of header
        LinkLabel developerLink = new LinkLabel
        {
            Text = "Developed by ShioDev",
            ForeColor = System.Drawing.Color.White,
            LinkColor = System.Drawing.Color.White,
            ActiveLinkColor = System.Drawing.Color.FromArgb(175, 185, 209), // Lighter color for hover
            VisitedLinkColor = System.Drawing.Color.White,
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            AutoSize = true
        };
        developerLink.LinkBehavior = LinkBehavior.HoverUnderline;
        developerLink.Click += DeveloperLink_Click;
        headerPanel.Controls.Add(developerLink);
        
        // Anchor the developer link to the right of the header
        developerLink.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
        developerLink.Location = new System.Drawing.Point(headerPanel.Width - developerLink.Width - 20, 40);
        
        // Create main container panel
        Panel mainPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(20),
            BackColor = System.Drawing.Color.FromArgb(20, 21, 30) // Dark background like Bitwarden
        };
        
        // Create left panel for controls
        Panel leftPanel = new Panel
        {
            Width = 220,
            Height = 350,
            Location = new System.Drawing.Point(20, 20),
            BackColor = System.Drawing.Color.FromArgb(31, 34, 52), // Dark panel like Bitwarden
            BorderStyle = BorderStyle.None
        };
        
        // Add shadow effect
        leftPanel.Paint += (sender, e) => {
            ControlPaint.DrawBorder(e.Graphics, leftPanel.ClientRectangle,
                System.Drawing.Color.FromArgb(41, 44, 62), 1, ButtonBorderStyle.Solid,
                System.Drawing.Color.FromArgb(41, 44, 62), 1, ButtonBorderStyle.Solid,
                System.Drawing.Color.FromArgb(51, 54, 72), 2, ButtonBorderStyle.Solid,
                System.Drawing.Color.FromArgb(51, 54, 72), 2, ButtonBorderStyle.Solid);
        };
        
        // Create right panel for status and display
        Panel rightPanel = new Panel
        {
            Width = 360,
            Height = 350,
            Location = new System.Drawing.Point(250, 20),
            BackColor = System.Drawing.Color.FromArgb(31, 34, 52), // Dark panel like Bitwarden
            BorderStyle = BorderStyle.None
        };
        
        // Add shadow effect
        rightPanel.Paint += (sender, e) => {
            ControlPaint.DrawBorder(e.Graphics, rightPanel.ClientRectangle,
                System.Drawing.Color.FromArgb(41, 44, 62), 1, ButtonBorderStyle.Solid,
                System.Drawing.Color.FromArgb(41, 44, 62), 1, ButtonBorderStyle.Solid,
                System.Drawing.Color.FromArgb(51, 54, 72), 2, ButtonBorderStyle.Solid,
                System.Drawing.Color.FromArgb(51, 54, 72), 2, ButtonBorderStyle.Solid);
        };
        
        // Mode selection group box
        GroupBox modeGroupBox = new GroupBox
        {
            Text = "Input Mode",
            Location = new System.Drawing.Point(15, 15),
            Size = new System.Drawing.Size(190, 80),
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            ForeColor = System.Drawing.Color.White,
            BackColor = System.Drawing.Color.FromArgb(31, 34, 52) // Dark background like Bitwarden
        };
        
        // Create radio buttons for mode selection
        RadioButton commandModeRadio = new RadioButton
        {
            Text = "Command Mode",
            Checked = true,
            AutoSize = true,
            Location = new System.Drawing.Point(15, 25),
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            Cursor = Cursors.Hand,
            ForeColor = System.Drawing.Color.White,
            BackColor = System.Drawing.Color.FromArgb(31, 34, 52) // Dark background like Bitwarden
        };
        
        RadioButton stringModeRadio = new RadioButton
        {
            Text = "String Mode",
            AutoSize = true,
            Location = new System.Drawing.Point(15, 50),
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            Cursor = Cursors.Hand,
            ForeColor = System.Drawing.Color.White,
            BackColor = System.Drawing.Color.FromArgb(31, 34, 52) // Dark background like Bitwarden
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
            BackColor = System.Drawing.Color.FromArgb(84, 130, 210), // Bitwarden accent blue
            ForeColor = System.Drawing.Color.White,
            Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        toggleButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(84, 130, 210); // Match the background
        toggleButton.Click += (sender, e) => ToggleRecognition(toggleButton);
        leftPanel.Controls.Add(toggleButton);
        
        // Add a clear text button
        Button clearTextButton = new Button
        {
            Text = "Clear Text",
            Size = new System.Drawing.Size(190, 40),
            Location = new System.Drawing.Point(15, 165),
            FlatStyle = FlatStyle.Flat,
            BackColor = System.Drawing.Color.FromArgb(41, 44, 62), // Darker button
            ForeColor = System.Drawing.Color.White,
            Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular),
            Cursor = Cursors.Hand
        };
        clearTextButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(51, 54, 72);
        
        clearTextButton.Click += (sender, e) => 
        {
            typedText.Clear();
            if (typedTextBox != null)
            {
                typedTextBox.Clear();
            }
        };
        leftPanel.Controls.Add(clearTextButton);
        
        /* Add a copy text button
        Button copyTextButton = new Button
        {
            Text = "Copy Text",
            Size = new System.Drawing.Size(190, 40),
            Location = new System.Drawing.Point(15, 165 + 45), // Position below clear button
            FlatStyle = FlatStyle.Flat,
            BackColor = System.Drawing.Color.FromArgb(41, 44, 62), // Darker button
            ForeColor = System.Drawing.Color.White,
            Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular),
            Cursor = Cursors.Hand,
            Visible = false // Hide the Copy Text button
        };
        copyTextButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(51, 54, 72);
        
        copyTextButton.Click += (sender, e) => 
        {
            if (typedTextBox != null && !string.IsNullOrEmpty(typedTextBox.Text))
            {
                Clipboard.SetText(typedTextBox.Text);
                UpdateStatusLabel("Text copied to clipboard");
            }
        };
        leftPanel.Controls.Add(copyTextButton);
        */

        // Add settings panel for speech recognition
        GroupBox settingsGroupBox = new GroupBox
        {
            Text = "Recognition Settings",
            Location = new System.Drawing.Point(15, 215), // Adjust position for new copy button
            Size = new System.Drawing.Size(190, 125), // Increase height to accommodate overlay checkbox
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            ForeColor = System.Drawing.Color.White,
            BackColor = System.Drawing.Color.FromArgb(31, 34, 52) // Dark background like Bitwarden
        };
        
        // Create label instead of checkbox for better styling
        useWhisperLabel = new Label
        {
            Text = "☐ Use AI Recognition",  // Unchecked box character
            AutoSize = true,
            Location = new System.Drawing.Point(15, 25),
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            Cursor = Cursors.Hand,
            ForeColor = System.Drawing.Color.White,
            BackColor = System.Drawing.Color.FromArgb(31, 34, 52)
        };
        
        // Set initial state
        if (useWhisperForString)
        {
            useWhisperLabel.Text = "☑ Use AI Recognition";  // Checked box character
        }
        
        // Add click handler to toggle
        useWhisperLabel.Click += (sender, e) => 
        {
            useWhisperForString = !useWhisperForString;
            useWhisperLabel.Text = useWhisperForString ? 
                "☑ Use AI Recognition" : 
                "☐ Use AI Recognition";
        };
        
        // Add Whisper model info as a simple label
        Label modelLabel = new Label
        {
            Text = "Whisper Model: Medium",
            AutoSize = true,
            Location = new System.Drawing.Point(15, 45),
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            ForeColor = System.Drawing.Color.FromArgb(150, 165, 195) // Medium gray-blue for model text
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
        settingsGroupBox.Controls.Add(useWhisperLabel);
        settingsGroupBox.Controls.Add(modelLabel);
        settingsGroupBox.Controls.Add(resetModelButton);
        
        leftPanel.Controls.Add(settingsGroupBox);
        
        // Create a checkbox for enabling the overlay
        enableOverlayCheckBox = new CheckBox
        {
            Text = "Enable floating button",
            AutoSize = true,
            Location = new Point(15, 75),
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            ForeColor = System.Drawing.Color.White,
            BackColor = System.Drawing.Color.FromArgb(31, 34, 52),
            Cursor = Cursors.Hand
        };
        enableOverlayCheckBox.CheckedChanged += EnableOverlayCheckBox_CheckedChanged;
        settingsGroupBox.Controls.Add(enableOverlayCheckBox);
        
        // Status panel
        Panel statusPanel = new Panel
        {
            Width = 340,
            Height = 70,
            Location = new System.Drawing.Point(10, 10),
            BackColor = System.Drawing.Color.FromArgb(25, 27, 38), // Slightly lighter than main background
            BorderStyle = BorderStyle.FixedSingle
        };
        
        // Create a status label
        Label statusLabel = new Label
        {
            Text = "Whisper ready with medium model - start listening",
            AutoSize = false,
            Width = 320,
            Height = 25,
            Location = new System.Drawing.Point(10, 10),
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            ForeColor = System.Drawing.Color.FromArgb(123, 162, 217) // Lighter blue for status text
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
            ForeColor = System.Drawing.Color.FromArgb(123, 162, 217) // Match the status text color
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
            ForeColor = System.Drawing.Color.FromArgb(175, 185, 209) // Light gray for text
        };
        rightPanel.Controls.Add(recognizedLabel);
        
        // Create speaking indicator (initially hidden)
        speakingNowLabel = new Label
        {
            Text = "SPEAK NOW - WILL PAUSE AFTER 1.5s SILENCE",
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
            BackColor = System.Drawing.Color.FromArgb(25, 27, 38), // Slightly lighter than main background
            ForeColor = System.Drawing.Color.White,
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
            BackColor = System.Drawing.Color.FromArgb(31, 34, 52), // Dark panel like Bitwarden
            BorderStyle = BorderStyle.FixedSingle
        };
        
        // Create instructions header
        Label instructionsHeader = new Label
        {
            Text = "Instructions",
            AutoSize = true,
            Location = new System.Drawing.Point(10, 8),
            Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold),
            ForeColor = System.Drawing.Color.FromArgb(84, 130, 210) // Bitwarden accent blue
        };
        instructionsPanel.Controls.Add(instructionsHeader);
        
        // Create instructions label
        Label instructionsLabel = new Label
        {
            Text = "• Command Mode: Say commands like \"up\", \"down\", \"enter\", \"control c\" for copy.\n" +
                  "• String Mode: Speak naturally and the system will automatically stop after 1.5 seconds of silence.\n" +
                  "• Say \"command option\" or \"string option\" to switch between modes at any time.\n" +
                  "• Recording indicator: ● Gray = Not recording, ● Blue = Recording in progress.\n" +
                  "• You can scroll through recognized text and copy it using the \"Copy Text\" button or right-click menu.",
            AutoSize = true,
            Location = new System.Drawing.Point(10, 30),
            Size = new System.Drawing.Size(580, 0),
            Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular),
            ForeColor = System.Drawing.Color.FromArgb(175, 185, 209) // Light gray for text
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
        
        // Update the overlay to show the current mode
        if (overlayForm != null && !overlayForm.IsDisposed && overlayEnabled)
        {
            overlayForm.UpdateButtonState(isListening, mode);
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
                UpdateRecordingStatus(true); // Show recording status for both modes
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
                            UpdateStatusLabel("Listening for speech - will pause after 1.5s of silence...");
                            
                            // Continuous recording with automatic silence detection
                            while (!cancellationTokenSource.Token.IsCancellationRequested)
                            {
                                // Show recording is active
                                UpdateRecordingStatus(true); 
                                
                                // Start recording session - this will stop automatically after silence detection
                                await audioCapture.StartRecordingAsync(cancellationTokenSource.Token);
                                
                                // Only hide the indicator if we're done listening entirely
                                if (cancellationTokenSource.Token.IsCancellationRequested)
                                {
                                    UpdateRecordingStatus(false);
                                }
                                
                                // Allow a small delay between recording sessions
                                await Task.Delay(300, cancellationTokenSource.Token);
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
            button.BackColor = System.Drawing.Color.FromArgb(84, 130, 210); // Use Bitwarden blue
            
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
                // Update the recording status indicator color
                recordingStatusLabel.ForeColor = isRecording 
                    ? System.Drawing.Color.FromArgb(84, 180, 210) // Bright blue for recording instead of green
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
                    speakingNowLabel.BackColor = System.Drawing.Color.FromArgb(84, 130, 210); // Bitwarden blue
                    speakingNowLabel.Text = "● SPEAK NOW - WILL PAUSE AFTER 1.5s SILENCE ●";
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
            
            // Also update overlay if available and not disposed
            try
            {
                if (overlayForm != null && !overlayForm.IsDisposed && overlayEnabled)
                {
                    overlayForm.UpdateButtonState(isRecording, currentMode);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating overlay: {ex.Message}");
            }
        });
    }

    private void InitializeFlashTimer()
    {
        flashTimer = new System.Windows.Forms.Timer();
        flashTimer.Interval = 500; // Flash every half second
        flashTimer.Tick += (sender, e) => 
        {
            if (speakingNowLabel != null && speakingNowLabel.Visible)
            {
                // Toggle between two shades of blue
                if (speakingNowLabel.BackColor == System.Drawing.Color.FromArgb(84, 130, 210))
                {
                    // Darker blue
                    speakingNowLabel.BackColor = System.Drawing.Color.FromArgb(64, 110, 190);
                }
                else
                {
                    // Back to normal blue
                    speakingNowLabel.BackColor = System.Drawing.Color.FromArgb(84, 130, 210);
                }
            }
        };
        flashTimer.Start();
    }

    // Clean up resources
    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        // If overlay is enabled and user is trying to close with the X button
        if (overlayEnabled && e.CloseReason == CloseReason.UserClosing)
        {
            // Cancel the close
            e.Cancel = true;
            
            // Hide the form instead
            this.Hide();
            mainWindowShown = false;
            
            // Show notification to the user
            if (overlayForm != null && !overlayForm.IsDisposed)
            {
                UpdateStatusLabel("Main window hidden. Use the floating button to control recording.");
            }
        }
        else
        {
            // Perform normal closing operations
            // Stop listening if active
            if (isListening)
            {
                StopListening();
            }
            
            // Cleanup resources
            commandRecognizer?.Dispose();
            dictationRecognizer?.Dispose();
            cancellationTokenSource?.Cancel();
            
            // Stop flash timer if running
            if (flashTimer != null)
            {
                flashTimer.Stop();
                flashTimer.Dispose();
                flashTimer = null;
            }
            
            // Clean up overlay
            try
            {
                if (overlayForm != null)
                {
                    if (!overlayForm.IsDisposed)
                    {
                        overlayForm.Hide();
                        overlayForm.Close();
                    }
                    overlayForm.Dispose();
                    overlayForm = null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error disposing overlay: {ex.Message}");
            }
            
            base.OnFormClosing(e);
        }
    }

    private void DeveloperLink_Click(object sender, EventArgs e)
    {
        // Show the about box with app information
        using (AboutBox aboutBox = new AboutBox())
        {
            aboutBox.ShowDialog(this);
        }
    }
    
    // Custom AboutBox form for displaying app information
    private class AboutBox : Form
    {
        private string appVersion;
        private string buildNumber;
        
        public AboutBox()
        {
            // Get version info from the parent form
            if (Owner is Form1 parentForm)
            {
                appVersion = parentForm.appVersion;
                buildNumber = parentForm.buildNumber;
            }
            else
            {
                // Default values if parent form not accessible
                appVersion = "1.0.2";
                buildNumber = "1002";
            }
            
            InitializeAboutBox();
        }
        
        private void InitializeAboutBox()
        {
            // Configure the form
            this.Text = "About";
            this.Size = new Size(450, 360);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.ShowInTaskbar = false;
            this.BackColor = System.Drawing.Color.FromArgb(20, 21, 30); // Dark background like Bitwarden
            
            // Use the same icon as the main form
            try
            {
                this.Icon = this.Owner.Icon;
            }
            catch
            {
                // Fallback silently
            }
            
            // Load app logo
            System.Drawing.Image logoImage;
            try
            {
                string iconPath = Path.Combine(Application.StartupPath, "icon", "ai-voice-keyboard.png");
                if (File.Exists(iconPath))
                {
                    logoImage = Image.FromFile(iconPath);
                }
                else
                {
                    // Try relative path
                    iconPath = Path.Combine("icon", "ai-voice-keyboard.png");
                    if (File.Exists(iconPath))
                    {
                        logoImage = Image.FromFile(iconPath);
                    }
                    else
                    {
                        // Fallback to system icon
                        logoImage = System.Drawing.SystemIcons.Application.ToBitmap();
                    }
                }
            }
            catch
            {
                // Fallback to system icon on any error
                logoImage = System.Drawing.SystemIcons.Application.ToBitmap();
            }
            
            // App logo/icon
            PictureBox logoBox = new PictureBox
            {
                Size = new Size(64, 64),
                Location = new Point(30, 30),
                SizeMode = PictureBoxSizeMode.StretchImage,
                Image = logoImage
            };
            
            // App title
            Label titleLabel = new Label
            {
                Text = "AI Voice Keyboard",
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = System.Drawing.Color.White,
                Location = new Point(110, 30),
                AutoSize = true
            };
            
            // Version info
            Label versionLabel = new Label
            {
                Text = $"Version: v{appVersion} (Build {buildNumber})",
                Font = new Font("Segoe UI", 10),
                ForeColor = System.Drawing.Color.FromArgb(175, 185, 209), // Light gray for text
                Location = new Point(110, 65),
                AutoSize = true
            };
            
            // Developer info
            Label devLabel = new Label
            {
                Text = "Developer:",
                Font = new Font("Segoe UI", 10),
                ForeColor = System.Drawing.Color.FromArgb(175, 185, 209), // Light gray for text
                Location = new Point(30, 120),
                AutoSize = true
            };
            
            // Developer link
            LinkLabel devLinkLabel = new LinkLabel
            {
                Text = "ShioDev",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = System.Drawing.Color.White,
                LinkColor = System.Drawing.Color.White,
                ActiveLinkColor = System.Drawing.Color.FromArgb(175, 185, 209), // Light blue hover
                VisitedLinkColor = System.Drawing.Color.White,
                Location = new Point(110, 120),
                AutoSize = true
            };
            devLinkLabel.LinkClicked += (s, e) => { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = "https://hello.shiodev.com", UseShellExecute = true }); };
            
            // Developer website
            Label devUrlLabel = new Label
            {
                Text = "https://hello.shiodev.com",
                Font = new Font("Segoe UI", 9),
                ForeColor = System.Drawing.Color.FromArgb(120, 130, 154), // Muted text
                Location = new Point(110, 145),
                AutoSize = true
            };
            
            // Foundation info
            Label foundationLabel = new Label
            {
                Text = "Powered by:",
                Font = new Font("Segoe UI", 10),
                ForeColor = System.Drawing.Color.FromArgb(175, 185, 209), // Light gray for text
                Location = new Point(30, 180),
                AutoSize = true
            };
            
            // Foundation link
            LinkLabel foundationLinkLabel = new LinkLabel
            {
                Text = "Beaver Foundation",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = System.Drawing.Color.White,
                LinkColor = System.Drawing.Color.White,
                ActiveLinkColor = System.Drawing.Color.FromArgb(175, 185, 209), // Light blue hover
                VisitedLinkColor = System.Drawing.Color.White,
                Location = new Point(110, 180),
                AutoSize = true
            };
            foundationLinkLabel.LinkClicked += (s, e) => { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = "https://beaver.foundation", UseShellExecute = true }); };
            
            // Foundation website
            Label foundationUrlLabel = new Label
            {
                Text = "https://beaver.foundation",
                Font = new Font("Segoe UI", 9),
                ForeColor = System.Drawing.Color.FromArgb(120, 130, 154), // Muted text
                Location = new Point(110, 205),
                AutoSize = true
            };
            
            // License info
            Label licenseLabel = new Label
            {
                Text = "License:",
                Font = new Font("Segoe UI", 10),
                ForeColor = System.Drawing.Color.FromArgb(175, 185, 209), // Light gray for text
                Location = new Point(30, 240),
                AutoSize = true
            };
            
            // License link
            LinkLabel licenseLinkLabel = new LinkLabel
            {
                Text = "GPL 3.0",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = System.Drawing.Color.White,
                LinkColor = System.Drawing.Color.White,
                ActiveLinkColor = System.Drawing.Color.FromArgb(175, 185, 209), // Light blue hover
                VisitedLinkColor = System.Drawing.Color.White,
                Location = new Point(110, 240),
                AutoSize = true
            };
            licenseLinkLabel.LinkClicked += (s, e) => { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = "https://www.gnu.org/licenses/gpl-3.0.en.html", UseShellExecute = true }); };
            
            // OK button
            Button okButton = new Button
            {
                Text = "OK",
                Size = new Size(80, 30),
                Location = new Point(this.Width / 2 - 40, 280),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(84, 130, 210), // Bitwarden accent blue
                ForeColor = System.Drawing.Color.White
            };
            okButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(84, 130, 210); // Match the button color
            okButton.Click += (s, e) => this.Close();
            
            // Add controls to the form
            this.Controls.Add(logoBox);
            this.Controls.Add(titleLabel);
            this.Controls.Add(versionLabel);
            this.Controls.Add(devLabel);
            this.Controls.Add(devLinkLabel);
            this.Controls.Add(devUrlLabel);
            this.Controls.Add(foundationLabel);
            this.Controls.Add(foundationLinkLabel);
            this.Controls.Add(foundationUrlLabel);
            this.Controls.Add(licenseLabel);
            this.Controls.Add(licenseLinkLabel);
            this.Controls.Add(okButton);
        }
    }

    private void EnableOverlayCheckBox_CheckedChanged(object? sender, EventArgs e)
    {
        if (enableOverlayCheckBox == null) return;
        
        overlayEnabled = enableOverlayCheckBox.Checked;
        
        try
        {
            if (overlayEnabled)
            {
                ShowOverlay();
                
                // Add notification to user about minimizing
                UpdateStatusLabel("You can now minimize or close the main window. The floating button will remain.");
            }
            else
            {
                HideOverlay();
                
                // Show the main window if it was hidden
                ShowMainWindow();
            }
        }
        catch (Exception ex)
        {
            UpdateStatusLabel($"Overlay error: {ex.Message}");
        }
    }
    
    private void ShowOverlay()
    {
        try
        {
            // Dispose old overlay if it exists
            if (overlayForm != null)
            {
                if (!overlayForm.IsDisposed)
                {
                    overlayForm.Hide();
                    overlayForm.Close();
                }
                overlayForm = null;
            }
            
            // Create a new overlay form
            overlayForm = new OverlayForm(this);
            overlayForm.UpdateButtonState(isListening, currentMode);
            
            // Show the form without setting this as owner to prevent minimizing with main window
            overlayForm.Show();
            
            // Ensure it's on top
            overlayForm.TopMost = true;
            overlayForm.BringToFront();
        }
        catch (Exception ex)
        {
            UpdateStatusLabel($"Error showing overlay: {ex.Message}");
        }
    }
    
    private void HideOverlay()
    {
        try
        {
            if (overlayForm != null)
            {
                if (!overlayForm.IsDisposed)
                {
                    overlayForm.Hide();
                    overlayForm.Close();
                }
                overlayForm = null;
            }
        }
        catch (Exception ex)
        {
            UpdateStatusLabel($"Error hiding overlay: {ex.Message}");
        }
    }
    
    public void ToggleRecordingFromOverlay()
    {
        try
        {
            // Find the toggle button in the controls
            Button? toggleButton = null;
            foreach (Control control in this.Controls)
            {
                if (control is Panel mainPanel)
                {
                    foreach (Control panelControl in mainPanel.Controls)
                    {
                        if (panelControl is Panel leftPanel)
                        {
                            foreach (Control leftControl in leftPanel.Controls)
                            {
                                if (leftControl is Button button && 
                                    (button.Text == "Start Listening" || button.Text == "Stop Listening"))
                                {
                                    toggleButton = button;
                                    break;
                                }
                            }
                        }
                        if (toggleButton != null) break;
                    }
                }
                if (toggleButton != null) break;
            }
            
            // Toggle the recording state using the found button or directly
            if (toggleButton != null && this.IsHandleCreated && !this.IsDisposed)
            {
                // Use BeginInvoke to safely invoke from another thread if needed
                this.BeginInvoke(new Action(() => ToggleRecognition(toggleButton)));
            }
            else
            {
                // Fallback to direct toggling if button not found
                if (isListening)
                {
                    StopListening();
                    isListening = false;
                    UpdateStatusLabel("Status: Not Listening");
                }
                else
                {
                    StartListening();
                    UpdateStatusLabel($"Listening in {currentMode} mode...");
                }
                
                // Update overlay button appearance manually
                if (overlayForm != null && !overlayForm.IsDisposed)
                {
                    overlayForm.UpdateButtonState(isListening, currentMode);
                }
            }
        }
        catch (Exception ex)
        {
            UpdateStatusLabel($"Error toggling recording: {ex.Message}");
        }
    }
    
    private void Form1_Resize(object sender, EventArgs e)
    {
        // When minimized and overlay is enabled, hide the window instead of minimizing
        if (this.WindowState == FormWindowState.Minimized && overlayEnabled)
        {
            this.Hide();
            mainWindowShown = false;
            
            // Make sure overlay is visible and properly displayed
            if (overlayForm != null && !overlayForm.IsDisposed)
            {
                // Reset owner to null to prevent overlay from minimizing with main form
                overlayForm.Owner = null;
                
                // Ensure it's visible and on top
                if (!overlayForm.Visible)
                {
                    overlayForm.Show();
                }
                overlayForm.TopMost = true;
                overlayForm.BringToFront();
                
                // Ensure the system knows the form is still active
                overlayForm.Update();
            }
        }
    }
    
    // Method to show the main window
    public void ShowMainWindow()
    {
        if (!mainWindowShown)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.BringToFront();
            mainWindowShown = true;
        }
    }
    
    // Add a proper exit method to truly close the application
    public void ExitApplication()
    {
        // Set a flag to truly close
        overlayEnabled = false;
        
        // Use Application.Exit to close all forms
        Application.Exit();
    }

    public void SwitchModeFromOverlay(InputMode mode)
    {
        SwitchMode(mode);
    }

    public void HideOverlayAndUntickCheckbox()
    {
        // Hide the overlay
        HideOverlay();
        
        // Untick the checkbox if it exists
        if (enableOverlayCheckBox != null && enableOverlayCheckBox.Checked)
        {
            enableOverlayCheckBox.Checked = false;
            overlayEnabled = false;
        }
    }
}

