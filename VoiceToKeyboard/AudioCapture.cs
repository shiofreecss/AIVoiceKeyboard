using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using NAudio.Wave;

namespace VoiceToKeyboard
{
    public class AudioCapture : IDisposable
    {
        private WaveInEvent? _waveIn;
        private MemoryStream? _memoryStream;
        private WaveFileWriter? _waveWriter;
        private bool _isRecording = false;
        private readonly SemaphoreSlim _recordingSemaphore = new SemaphoreSlim(1, 1);
        private float _lastVolume = 0;
        private bool _hasSpeech = false;
        private int _silenceCounter = 0;
        
        public event EventHandler<byte[]>? AudioCaptured;
        public event EventHandler<string>? StatusChanged;
        
        public bool IsRecording => _isRecording;
        
        // Settings for audio capture - adjusted for quicker word detection
        private const int SampleRate = 16000;         // 16kHz is optimal for Whisper
        private const int BitsPerSample = 16;         // 16-bit audio
        private const int Channels = 1;               // Mono
        private const int RecordingLengthMs = 2500;   // 2.5 seconds max recording before processing (reduced from 5s)
        private const float SilenceThreshold = 0.02f; // Threshold to detect speech
        private const int MinimumAudioLength = 500;   // Minimum length in milliseconds (reduced from 1000ms)
        private const int SilenceDetectionMs = 750;   // Detect silence after 750ms for quicker word detection
        
        public async Task StartRecordingAsync(CancellationToken cancellationToken)
        {
            await _recordingSemaphore.WaitAsync();
            
            try
            {
                if (_isRecording)
                    return;
                
                _isRecording = true;
                _hasSpeech = false;
                _silenceCounter = 0;
                
                RaiseStatusChanged("Starting audio capture...");
                
                // Initialize memory stream to store audio
                _memoryStream = new MemoryStream();
                
                // Create wave in device with format compatible with Whisper
                _waveIn = new WaveInEvent
                {
                    WaveFormat = new WaveFormat(SampleRate, BitsPerSample, Channels),
                    BufferMilliseconds = 50  // Reduced buffer for quicker processing
                };
                
                // Create wave writer with same format
                _waveWriter = new WaveFileWriter(_memoryStream, _waveIn.WaveFormat);
                
                // Setup event handlers
                _waveIn.DataAvailable += OnDataAvailable;
                _waveIn.RecordingStopped += OnRecordingStopped;
                
                // Start recording
                _waveIn.StartRecording();
                
                // Run a task to monitor for speech pause or max recording time
                _ = Task.Run(async () => 
                {
                    try 
                    {
                        int elapsedTime = 0;
                        bool hasFinishedNaturally = false;
                        
                        // Continue recording until we detect silence after speech
                        // or hit maximum recording length
                        while (elapsedTime < RecordingLengthMs && !hasFinishedNaturally && !cancellationToken.IsCancellationRequested)
                        {
                            await Task.Delay(50, cancellationToken); // Check more frequently (50ms instead of 100ms)
                            elapsedTime += 50;
                            
                            // If we've detected speech and now there's silence again
                            // Reduce silence duration needed to finish - 8 * 50ms = 400ms silence
                            if (_hasSpeech && _silenceCounter >= 8) 
                            {
                                hasFinishedNaturally = true;
                                RaiseStatusChanged("Word detected, processing...");
                            }
                        }
                        
                        // Stop recording if there was speech or we hit maximum length
                        if (_hasSpeech || elapsedTime >= MinimumAudioLength)
                        {
                            await StopRecordingAsync();
                        }
                        else
                        {
                            // No speech detected, just discard
                            RaiseStatusChanged("No speech detected, discarding...");
                            await StopRecordingAsync(false);
                        }
                    }
                    catch (TaskCanceledException)
                    {
                        // This is expected on cancellation
                    }
                }, cancellationToken);
            }
            catch (Exception ex)
            {
                RaiseStatusChanged($"Audio capture error: {ex.Message}");
                _isRecording = false;
                DisposeRecordingResources();
            }
            finally
            {
                _recordingSemaphore.Release();
            }
        }
        
        public async Task StopRecordingAsync(bool processAudio = true)
        {
            await _recordingSemaphore.WaitAsync();
            
            try
            {
                if (!_isRecording)
                    return;
                
                RaiseStatusChanged("Stopping audio capture...");
                
                // Stop the recording device
                _waveIn?.StopRecording();
                
                // Flush and close the wave writer
                _waveWriter?.Flush();
                
                if (_memoryStream != null && processAudio)
                {
                    try 
                    {
                        // Rewind the stream
                        _memoryStream.Position = 0;
                        
                        // Check if we have enough audio data
                        if (_memoryStream.Length > SampleRate * BitsPerSample * Channels / 8 * MinimumAudioLength / 1000)
                        {
                            RaiseStatusChanged($"Audio captured: {_memoryStream.Length} bytes");
                            
                            // Normalize the audio (boost quiet speech)
                            byte[] normalizedAudio = NormalizeAudio(_memoryStream.ToArray());
                            
                            // Raise event with the captured audio
                            AudioCaptured?.Invoke(this, normalizedAudio);
                        }
                        else
                        {
                            RaiseStatusChanged("Audio too short, ignoring");
                        }
                    }
                    catch (Exception ex)
                    {
                        RaiseStatusChanged($"Error processing audio data: {ex.Message}");
                    }
                }
                
                _isRecording = false;
                DisposeRecordingResources();
            }
            catch (Exception ex)
            {
                RaiseStatusChanged($"Error stopping recording: {ex.Message}");
            }
            finally
            {
                _recordingSemaphore.Release();
            }
        }
        
        private void OnDataAvailable(object? sender, WaveInEventArgs e)
        {
            try
            {
                // Calculate volume level to detect speech
                float volume = CalculateVolume(e.Buffer, e.BytesRecorded);
                
                // Update last volume
                _lastVolume = volume;
                
                // Check if this is speech or silence
                if (volume > SilenceThreshold)
                {
                    _hasSpeech = true;
                    _silenceCounter = 0;
                    // RaiseStatusChanged($"Speech detected: {volume:F3}");
                }
                else if (_hasSpeech)
                {
                    _silenceCounter++;
                }
                
                // Write audio data to wave writer
                if (_waveWriter != null && e.BytesRecorded > 0)
                {
                    _waveWriter.Write(e.Buffer, 0, e.BytesRecorded);
                }
            }
            catch (Exception ex)
            {
                RaiseStatusChanged($"Error writing audio data: {ex.Message}");
            }
        }
        
        private float CalculateVolume(byte[] buffer, int bytesRecorded)
        {
            float sum = 0;
            int samplesCount = bytesRecorded / 2; // 16-bit samples
            
            for (int i = 0; i < bytesRecorded; i += 2)
            {
                // Convert two bytes to a 16-bit sample
                short sample = (short)((buffer[i + 1] << 8) | buffer[i]);
                
                // Get absolute value and normalize to 0-1 range
                float normalized = Math.Abs(sample) / 32768f;
                sum += normalized;
            }
            
            return sum / samplesCount; // Average volume across samples
        }
        
        private byte[] NormalizeAudio(byte[] audioData)
        {
            // For 16-bit PCM WAV, the first 44 bytes are the header
            const int headerSize = 44;
            
            // If audio is too small, just return it
            if (audioData.Length <= headerSize)
                return audioData;
                
            // Extract the WAV header
            byte[] header = new byte[headerSize];
            Array.Copy(audioData, 0, header, 0, headerSize);
            
            // Get the audio samples (16-bit PCM)
            byte[] rawAudio = new byte[audioData.Length - headerSize];
            Array.Copy(audioData, headerSize, rawAudio, 0, rawAudio.Length);
            
            // Convert to short samples
            short[] samples = new short[rawAudio.Length / 2];
            for (int i = 0; i < samples.Length; i++)
            {
                samples[i] = (short)((rawAudio[i * 2 + 1] << 8) | rawAudio[i * 2]);
            }
            
            // Calculate max absolute sample value
            short maxSample = 0;
            foreach (short sample in samples)
            {
                maxSample = (short)Math.Max(maxSample, Math.Abs(sample));
            }
            
            // Calculate gain to apply
            float gain = maxSample < 16384 ? 32767f / maxSample : 1.0f;
            gain = Math.Min(gain, 4.0f); // Limit gain to avoid excessive amplification of quiet audio
            
            // Apply gain to samples
            for (int i = 0; i < samples.Length; i++)
            {
                float amplified = samples[i] * gain;
                
                // Clamp to short range
                if (amplified > 32767) amplified = 32767;
                if (amplified < -32768) amplified = -32768;
                
                samples[i] = (short)amplified;
            }
            
            // Convert back to bytes
            for (int i = 0; i < samples.Length; i++)
            {
                rawAudio[i * 2] = (byte)(samples[i] & 0xFF);
                rawAudio[i * 2 + 1] = (byte)((samples[i] >> 8) & 0xFF);
            }
            
            // Combine header and modified audio
            byte[] result = new byte[audioData.Length];
            Array.Copy(header, 0, result, 0, headerSize);
            Array.Copy(rawAudio, 0, result, headerSize, rawAudio.Length);
            
            return result;
        }
        
        private void OnRecordingStopped(object? sender, StoppedEventArgs e)
        {
            if (e.Exception != null)
            {
                RaiseStatusChanged($"Recording error: {e.Exception.Message}");
            }
        }
        
        private void DisposeRecordingResources()
        {
            if (_waveWriter != null)
            {
                try
                {
                    _waveWriter.Dispose();
                }
                catch { }
                _waveWriter = null;
            }
            
            if (_memoryStream != null)
            {
                try
                {
                    _memoryStream.Dispose();
                }
                catch { }
                _memoryStream = null;
            }
            
            if (_waveIn != null)
            {
                try
                {
                    _waveIn.Dispose();
                }
                catch { }
                _waveIn = null;
            }
        }
        
        private void RaiseStatusChanged(string status)
        {
            StatusChanged?.Invoke(this, status);
        }
        
        public void Dispose()
        {
            DisposeRecordingResources();
            _recordingSemaphore.Dispose();
        }
    }
} 