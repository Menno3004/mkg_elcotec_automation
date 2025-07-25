// 🎯 Part 2: RunHistoryManager Class
// Add this to a new file: Services/RunHistoryManager.cs

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Mkg_Elcotec_Automation.Models;

namespace Mkg_Elcotec_Automation.Services
{
    /// <summary>
    /// Manages saving/loading of automation runs
    /// </summary>
    public class RunHistoryManager
    {
        private readonly string _historyDirectory;
        private readonly string _historyIndexFile;
        private List<RunHistoryItem> _runHistory;

        public RunHistoryManager()
        {
            var baseAppDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MKG_Elcotec_Automation"
            );

            _historyDirectory = Path.Combine(baseAppDataPath, "RunHistory");
            _historyIndexFile = Path.Combine(_historyDirectory, "run_index.json");

            // Create both directories
            Directory.CreateDirectory(_historyDirectory);
            Directory.CreateDirectory(Path.Combine(baseAppDataPath, "DebugLogs"));

            LoadRunHistory();

            Console.WriteLine($"✅ RunHistory folder: {_historyDirectory}");
            Console.WriteLine($"✅ DebugLogs folder: {Path.Combine(baseAppDataPath, "DebugLogs")}");
        }

        /// <summary>
        /// Get all runs ordered by date (newest first)
        /// </summary>
        public List<RunHistoryItem> GetAllRuns()
        {
            try
            {
                // Reload from disk to ensure we have the latest data
                LoadRunHistory();

                // Return ordered by date (newest first) with current run at top if it exists
                var orderedRuns = _runHistory.OrderByDescending(r => r.IsCurrentRun)
                                           .ThenByDescending(r => r.StartTime)
                                           .ToList();

                Console.WriteLine($"📊 GetAllRuns returning {orderedRuns.Count} runs");

                return orderedRuns;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in GetAllRuns: {ex.Message}");
                return new List<RunHistoryItem>();
            }
        }
        public Guid CreateNewRun()
        {
            var newRun = new RunHistoryItem
            {
                RunId = Guid.NewGuid(),
                StartTime = DateTime.Now,
                Status = "Running",
                IsCurrentRun = true
            };

            // Mark all other runs as not current
            foreach (var run in _runHistory)
            {
                run.IsCurrentRun = false;
            }

            _runHistory.Add(newRun);
            SaveRunHistory();

            Console.WriteLine($"✅ Created new run: {newRun.RunId}");
            return newRun.RunId;
        }

        /// <summary>
        /// Get current run if any
        /// </summary>
        public RunHistoryItem GetCurrentRun()
        {
            return _runHistory.FirstOrDefault(r => r.IsCurrentRun);
        }
        /// <summary>
        /// Update run progress
        /// </summary>
        public void UpdateRunProgress(Guid runId, int emailsProcessed, int ordersFound,
                                    int quotesFound, int revisionsFound, int totalItems)
        {
            var run = _runHistory.FirstOrDefault(r => r.RunId == runId);
            if (run != null)
            {
                run.EmailsProcessed = emailsProcessed;
                run.OrdersFound = ordersFound;
                run.QuotesFound = quotesFound;
                run.RevisionsFound = revisionsFound;
                run.TotalItems = totalItems;
                SaveRunHistory();
            }
        }

        /// <summary>
        /// Complete a run
        /// </summary>
        public void CompleteRun(Guid runId, string status = "Completed", bool hasInjectionFailures = false, int totalFailures = 0, int totalDuplicates = 0)
        {
            var run = _runHistory.FirstOrDefault(r => r.RunId == runId);
            if (run != null)
            {
                run.Status = status;
                run.EndTime = DateTime.Now;
                run.IsCurrentRun = false;

                // Save injection failure status
                run.HasInjectionFailures = hasInjectionFailures;
                run.TotalFailuresAtCompletion = totalFailures;

                // 🎯 NEW: Save enhanced tab color status to RunHistoryItem as well
                if (run.Settings == null)
                    run.Settings = new Dictionary<string, object>();

                run.Settings["TotalErrors"] = totalFailures;
                run.Settings["TotalDuplicates"] = totalDuplicates;

                // Determine tab color status
                var tabColorStatus = totalFailures > 0 ? "Red" : (totalDuplicates > 0 ? "Yellow" : "Green");
                run.Settings["TabColorStatus"] = tabColorStatus;

                SaveRunHistory();

                Console.WriteLine($"✅ Completed run {runId} with status: {status}");
                Console.WriteLine($"   📊 Final counts preserved");
                Console.WriteLine($"   🎨 Tab color: {tabColorStatus} ({totalFailures} errors, {totalDuplicates} duplicates)");
            }
        }

        /// <summary>
        /// Load complete run data
        /// </summary>
        public AutomationRunData LoadRunData(Guid runId)
        {
            var runFile = Path.Combine(_historyDirectory, $"{runId}.json");
            if (File.Exists(runFile))
            {
                try
                {
                    var json = File.ReadAllText(runFile);
                    return JsonSerializer.Deserialize<AutomationRunData>(json);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading run data: {ex.Message}");
                }
            }
            return null;
        }

        /// <summary>
        /// Save complete run data
        /// </summary>
        public void SaveRunData(AutomationRunData runData)
        {
            try
            {
                // 🔧 FIX: Check for null runData and RunInfo
                if (runData == null)
                {
                    Console.WriteLine("❌ Cannot save run data - runData is null");
                    return;
                }

                if (runData.RunInfo == null)
                {
                    Console.WriteLine("❌ Cannot save run data - runData.RunInfo is null");
                    return;
                }

                var runFile = Path.Combine(_historyDirectory, $"{runData.RunInfo.RunId}.json");
                var options = new JsonSerializerOptions { WriteIndented = true };
                var json = JsonSerializer.Serialize(runData, options);
                File.WriteAllText(runFile, json);

                Console.WriteLine($"✅ Successfully saved run data for: {runData.RunInfo.RunId}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error saving run data: {ex.Message}");
                Console.WriteLine($"   Stack trace: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// Delete old runs (keep last 50)
        /// </summary>
        public void CleanupOldRuns()
        {
            var oldRuns = _runHistory.OrderByDescending(r => r.StartTime).Skip(50).ToList();
            foreach (var run in oldRuns)
            {
                try
                {
                    var runFile = Path.Combine(_historyDirectory, $"{run.RunId}.json");
                    if (File.Exists(runFile))
                        File.Delete(runFile);

                    _runHistory.Remove(run);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error cleaning up run {run.RunId}: {ex.Message}");
                }
            }
            SaveRunHistory();
        }
       
        #region Private Methods

        private void LoadRunHistory()
        {
            try
            {
                if (File.Exists(_historyIndexFile))
                {
                    var json = File.ReadAllText(_historyIndexFile);
                    _runHistory = JsonSerializer.Deserialize<List<RunHistoryItem>>(json) ?? new List<RunHistoryItem>();
                }
                else
                {
                    _runHistory = new List<RunHistoryItem>();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading run history: {ex.Message}");
                _runHistory = new List<RunHistoryItem>();
            }
        }

        private void SaveRunHistory()
        {
            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                var json = JsonSerializer.Serialize(_runHistory, options);
                File.WriteAllText(_historyIndexFile, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving run history: {ex.Message}");
            }
        }

        #endregion
    }
}