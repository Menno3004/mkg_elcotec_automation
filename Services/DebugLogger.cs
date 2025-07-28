using System;
using System.IO;
using System.Text;

namespace Mkg_Elcotec_Automation.Services
{
    /// <summary>
    /// Debug logger that creates ONE concise file per run - fits in Claude context window
    /// </summary>
    public static class DebugLogger
    {
        private static readonly string _debugDirectory;
        private static string _currentRunFile;
        private static StringBuilder _runBuffer = new StringBuilder();
        private static readonly int MAX_ENTRIES = 200; // Keep it small for context window

        static DebugLogger()
        {
            var baseAppDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MKG_Elcotec_Automation"
            );

            _debugDirectory = Path.Combine(baseAppDataPath, "DebugLogs");
            Directory.CreateDirectory(_debugDirectory);
        }

        public static void LogDebug(string message, string source = "System")
        {
            try
            {
                if (string.IsNullOrEmpty(_currentRunFile))
                {
                    StartNewRun("auto_started");
                }

                var timestamp = DateTime.Now.ToString("HH:mm:ss");
                var logEntry = $"[{timestamp}] [{source}] {message}";

                if (_runBuffer.ToString().Split('\n').Length < MAX_ENTRIES)
                {
                    _runBuffer.AppendLine(logEntry);
                }

                Console.WriteLine(logEntry);

                // Write immediately to file
                try
                {
                    File.AppendAllText(_currentRunFile, logEntry + Environment.NewLine);
                }
                catch { /* ignore file errors */ }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error logging: {ex.Message}");
            }
        }

        /// <summary>
        /// Log error with stack trace
        /// </summary>
        public static void LogError(string message, Exception ex = null, string source = "System")
        {
            try
            {
                // ✅ FIX: Auto-initialize if not started
                if (string.IsNullOrEmpty(_currentRunFile))
                {
                    StartNewRun("auto_started");
                }

                var timestamp = DateTime.Now.ToString("HH:mm:ss");
                var logEntry = $"[{timestamp}] [ERROR] [{source}] {message}";

                if (ex != null)
                {
                    logEntry += $" | Exception: {ex.Message}";
                }

                if (_runBuffer.ToString().Split('\n').Length < MAX_ENTRIES)
                {
                    _runBuffer.AppendLine(logEntry);
                }

                Console.WriteLine(logEntry);

                // Write immediately to file
                try
                {
                    File.AppendAllText(_currentRunFile, logEntry + Environment.NewLine);
                }
                catch { /* ignore file errors */ }
            }
            catch (Exception logEx)
            {
                Console.WriteLine($"❌ Error logging error: {logEx.Message}");
            }
        }

        /// <summary>
        /// Start a new run - creates a single file for this entire run
        /// </summary>
        public static void StartNewRun(string runName = "mkg_run")
        {
            // Generate unique timestamp with milliseconds to avoid conflicts
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss-fff");
            _currentRunFile = Path.Combine(_debugDirectory, $"{runName}_{timestamp}.log");

            // Clear buffer for new run
            _runBuffer.Clear();

            var header = $"=== MKG RUN DEBUG LOG ===" + Environment.NewLine +
                        $"Started: {DateTime.Now:yyyy-MM-dd HH:mm:ss}" + Environment.NewLine +
                        $"User: {Environment.UserName}" + Environment.NewLine +
                        $"======================================" + Environment.NewLine + Environment.NewLine;

            _runBuffer.AppendLine($"=== MKG RUN DEBUG LOG ===");
            _runBuffer.AppendLine($"Started: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            _runBuffer.AppendLine($"User: {Environment.UserName}");
            _runBuffer.AppendLine($"======================================");
            _runBuffer.AppendLine();

            // Write header to file immediately
            try
            {
                File.WriteAllText(_currentRunFile, header);
                Console.WriteLine($"✅ New run started: {_currentRunFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error creating run file: {ex.Message}");
            }
        }

        /// <summary>
        /// Save current run data to file
        /// </summary>
        public static void SaveCurrentRun()
        {
            // Not needed - we write immediately now
        }

        /// <summary>
        /// Get the current run's debug content (for immediate access)
        /// </summary>
        public static string GetCurrentRunLog()
        {
            return _runBuffer.ToString();
        }

        /// <summary>
        /// Get debug logs directory path
        /// </summary>
        public static string GetDebugDirectory()
        {
            return _debugDirectory;
        }

        /// <summary>
        /// LEGACY COMPATIBILITY: Create session log (redirects to current run file)
        /// </summary>
        public static string CreateSessionLog(string sessionName)
        {
            if (string.IsNullOrEmpty(_currentRunFile))
            {
                StartNewRun(sessionName);
            }
            return _currentRunFile;
        }

        /// <summary>
        /// LEGACY COMPATIBILITY: Get current log file
        /// </summary>
        public static string GetCurrentLogFile()
        {
            if (string.IsNullOrEmpty(_currentRunFile))
            {
                StartNewRun("default");
            }
            return _currentRunFile;
        }

        /// <summary>
        /// Clean up old log files (keep last 10 runs)
        /// </summary>
        public static void CleanupOldLogs()
        {
            try
            {
                var logFiles = Directory.GetFiles(_debugDirectory, "*.log");
                if (logFiles.Length > 10)
                {
                    Array.Sort(logFiles, (x, y) => File.GetCreationTime(x).CompareTo(File.GetCreationTime(y)));
                    for (int i = 0; i < logFiles.Length - 10; i++)
                    {
                        File.Delete(logFiles[i]);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error cleaning up logs: {ex.Message}");
            }
        }
    }
}