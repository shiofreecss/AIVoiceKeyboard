using System;
using System.IO;
using System.Threading.Tasks;
using Whisper.net;
using System.Threading;

namespace VoiceToKeyboard
{
    public class WhisperSpeechRecognition : IDisposable
    {
        private WhisperFactory? _whisperFactory;
        private WhisperProcessor? _processor;
        private bool _isReady = false;
        private string _modelFileName;
        
        public event EventHandler<string>? TextRecognized;
        public event EventHandler<string>? StatusChanged;
        
        public bool IsReady => _isReady;
        
        public WhisperSpeechRecognition(string modelFileName = "ggml-base.bin")
        {
            _modelFileName = modelFileName;
            
            // Get the full path to the model file, starting with the executable directory
            string exePath = AppDomain.CurrentDomain.BaseDirectory;
            _modelFileName = Path.Combine(exePath, modelFileName);
            
            RaiseStatusChanged($"Whisper model path: {_modelFileName}");
            
            // Check models directory
            CheckAndCreateModelsDirectory();
        }
        
        private void CheckAndCreateModelsDirectory()
        {
            try
            {
                // Get the models directory (same as executable directory)
                string exePath = AppDomain.CurrentDomain.BaseDirectory;
                string modelsDir = Path.Combine(exePath, "models");
                
                // Create models directory if it doesn't exist
                if (!Directory.Exists(modelsDir))
                {
                    Directory.CreateDirectory(modelsDir);
                    RaiseStatusChanged($"Created models directory at {modelsDir}");
                }
            }
            catch (Exception ex)
            {
                RaiseStatusChanged($"Error creating models directory: {ex.Message}");
            }
        }
        
        public async Task InitializeAsync()
        {
            try
            {
                RaiseStatusChanged("Starting Whisper initialization...");
                
                // Check if model exists
                if (!File.Exists(_modelFileName))
                {
                    string modelName = Path.GetFileName(_modelFileName);
                    RaiseStatusChanged($"Model file not found: {modelName}");
                    
                    // Try to find it in the models directory
                    string modelsDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "models");
                    string modelInModelsDir = Path.Combine(modelsDir, modelName);
                    
                    if (File.Exists(modelInModelsDir))
                    {
                        _modelFileName = modelInModelsDir;
                        RaiseStatusChanged($"Found model in models directory: {_modelFileName}");
                    }
                    else
                    {
                        // Provide detailed download instructions
                        RaiseStatusChanged($"Please download the model file and place it in the executable directory or 'models' folder");
                        
                        // Provide download links based on model size
                        if (modelName.Contains("tiny"))
                        {
                            RaiseStatusChanged("Download URL: https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-tiny.bin (75 MB)");
                        }
                        else if (modelName.Contains("base"))
                        {
                            RaiseStatusChanged("Download URL: https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-base.bin (142 MB)");
                        }
                        else if (modelName.Contains("small"))
                        {
                            RaiseStatusChanged("Download URL: https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-small.bin (466 MB)");
                        }
                        else if (modelName.Contains("medium"))
                        {
                            RaiseStatusChanged("Download URL: https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-medium.bin (1.5 GB)");
                        }
                        else if (modelName.Contains("large-v3-turbo"))
                        {
                            RaiseStatusChanged("Download URL: https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-large-v3-turbo.bin (1.5 GB)");
                        }
                        else if (modelName.Contains("large-v3"))
                        {
                            RaiseStatusChanged("Download URL: https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-large-v3.bin (2.9 GB)");
                        }
                        
                        RaiseStatusChanged("After downloading, restart the application");
                        throw new FileNotFoundException($"Whisper model file not found: {_modelFileName}");
                    }
                }
                
                RaiseStatusChanged($"Loading Whisper model...");
                
                try
                {
                    _whisperFactory = WhisperFactory.FromPath(_modelFileName);
                }
                catch (Exception modelEx)
                {
                    RaiseStatusChanged($"Error loading model: {modelEx.Message}");
                    RaiseStatusChanged("Make sure the model file is not corrupted and is a valid Whisper model");
                    throw;
                }
                
                // Create the processor with English language and optimized for real-time word processing
                RaiseStatusChanged("Creating speech processor for real-time word detection...");
                _processor = _whisperFactory.CreateBuilder()
                    .WithLanguage("en")
                    .Build();
                
                _isReady = true;
                RaiseStatusChanged("Whisper initialization complete - ready for real-time word processing");
            }
            catch (Exception ex)
            {
                _isReady = false;
                RaiseStatusChanged($"Whisper initialization failed: {ex.Message}");
                throw;
            }
        }
        
        public async Task<string> ProcessAudioAsync(byte[] audioData)
        {
            if (!_isReady || _processor == null)
                throw new InvalidOperationException("Whisper is not initialized");
                
            try
            {
                RaiseStatusChanged("Processing audio for words...");
                
                // Convert bytes to stream - ensure WAV format
                using var memoryStream = new MemoryStream(audioData);
                
                // Process with Whisper
                var sb = new System.Text.StringBuilder();
                bool foundText = false;
                
                try 
                {
                    // Process with a timeout for better real-time performance
                    using var tokenSource = new CancellationTokenSource();
                    tokenSource.CancelAfter(2000); // 2 seconds max processing time
                    
                    await foreach (var segment in _processor.ProcessAsync(memoryStream, tokenSource.Token))
                    {
                        // Skip segments with [BLANK_AUDIO] marker or empty text
                        if (segment.Text.Contains("[BLANK_AUDIO]") || string.IsNullOrWhiteSpace(segment.Text))
                        {
                            continue;
                        }
                        
                        // Process each word to get real-time words
                        string segmentText = segment.Text.Trim();
                        
                        // Basic word splitting for immediate feedback
                        // (Note: Whisper doesn't always provide perfect word boundaries)
                        string[] words = segmentText.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                        if (words.Length > 0)
                        {
                            foundText = true;
                            sb.Append(segmentText + " ");
                            
                            // Report progress on each segment for real-time feedback
                            RaiseStatusChanged($"Words detected: {segmentText}");
                        }
                    }
                }
                catch (OperationCanceledException)
                {
                    // This is expected if processing takes too long
                    RaiseStatusChanged("Processing timed out, returning partial results");
                }
                catch (Exception ex)
                {
                    RaiseStatusChanged($"Whisper processing error: {ex.Message}");
                    // Continue and return whatever we got
                }
                
                string result = sb.ToString().Trim();
                
                if (!foundText || string.IsNullOrWhiteSpace(result))
                {
                    RaiseStatusChanged("No words detected");
                    return string.Empty;
                }
                
                RaiseStatusChanged($"Words recognized: {result}");
                return result;
            }
            catch (Exception ex)
            {
                RaiseStatusChanged($"Processing error: {ex.Message}");
                throw;
            }
        }
        
        private void RaiseStatusChanged(string status)
        {
            StatusChanged?.Invoke(this, status);
        }
        
        public void Dispose()
        {
            _processor?.Dispose();
            _whisperFactory?.Dispose();
            GC.SuppressFinalize(this);
        }
    }
} 