using Mkg_Elcotec_Automation.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Mkg_Elcotec_Automation.Services
{
    /// <summary>
    /// Enhanced UI/UX Progress System - COMPLETE FIXED VERSION
    /// 🎯 FIXES: Debug console overwriting, statistics panel persistence
    /// </summary>
    public class EnhancedProgressManager
    {
        #region Private Fields

        private readonly Form _parentForm;
        private readonly TabControl _tabControl;
        private readonly ToolStripProgressBar _progressBar;
        private readonly ToolStripLabel _statusLabel;
        private readonly ToolStripLabel _progressLabel;

        // Clean floating stats form
        private Form _floatingStatsForm;
        private Label _floatingTimeElapsedLabel;
        private Label _floatingCurrentOperationLabel;
        private ToolTip _tooltip;

        // 🎯 NEW: Control flags for proper panel management
        private bool _isProcessingComplete = false;
        private bool _preventAutoHide = false;

        // Statistics Labels (Grid Layout)
        private Label _headerOrdersLabel;      // H.Orders
        private Label _lineOrdersLabel;        // L.Orders  
        private Label _totalEmailsLabel;       // T.Emails
        private Label _skippedEmailsLabel;     // S.Emails
        private Label _duplicateEmailsLabel;   // D.Emails
        private Label _headerQuotesLabel;      // H.Quotes
        private Label _lineQuotesLabel;        // L.Quotes
        private Label _injectionErrorsLabel;   // I.Errors
        private Label _businessErrorsLabel;    // B.Errors
        private Label _duplicateErrorsLabel;   // D.Errors
        private Label _headerRevisionsLabel;   // H.Revisions
        private Label _lineRevisionsLabel;     // L.Revisions

        // Statistics Counters (Clean Set)
        private int _headerOrders = 0;
        private int _lineOrders = 0;

        private int _headerRevisions = 0;
        private int _lineRevisions = 0;

        private int _headerQuotes = 0;
        private int _lineQuotes = 0;

        //EMAIL
        private int _totalEmails = 0;
        private int _skippedEmails = 0;
        private int _duplicateEmails = 0;

        //MKG
        private int _injectionErrors = 0;
        private int _businessErrors = 0;
        private int _duplicateErrors = 0;


        // Operation Control
        private DateTime _operationStartTime;
        private int _currentStep;
        private int _totalSteps;
        private string _currentOperation = "Initializing...";
        private Timer _updateTimer;

        private int _emailsProcessed = 0;
        private int _ordersFound = 0;
        private int _quotesFound = 0;
        private int _revisionsFound = 0;
        private string _lastOperation = "Initializing...";

        public int EmailsProcessed => _emailsProcessed;
        public int OrdersFound => _ordersFound;
        public int QuotesFound => _quotesFound;
        public int RevisionsFound => _revisionsFound;
        public string LastOperation => _lastOperation;
        public int GetInjectionErrors() => _injectionErrors;
        public int GetBusinessErrors() => _businessErrors;
        public int GetDuplicateErrors() => _duplicateErrors;

        #endregion

        #region Events

        public enum LogLevel { Info, Warning, Error, Success, Debug }

        public class LogEntryEventArgs : EventArgs
        {
            public string Message { get; set; }
            public LogLevel Level { get; set; }
            public DateTime Timestamp { get; set; } = DateTime.Now;
        }

        public event EventHandler<LogEntryEventArgs> LogEntryAdded;

        #endregion

        #region Constructor

        public EnhancedProgressManager(
            Form parentForm,
            TabControl tabControl,
            ToolStripProgressBar progressBar,
            ToolStripLabel statusLabel,
            ToolStripLabel progressLabel)
        {
            _parentForm = parentForm ?? throw new ArgumentNullException(nameof(parentForm));
            _tabControl = tabControl ?? throw new ArgumentNullException(nameof(tabControl));
            _progressBar = progressBar ?? throw new ArgumentNullException(nameof(progressBar));
            _statusLabel = statusLabel ?? throw new ArgumentNullException(nameof(statusLabel));
            _progressLabel = progressLabel ?? throw new ArgumentNullException(nameof(progressLabel));

            InitializeUpdateTimer();
            Console.WriteLine("✅ EnhancedProgressManager initialized (Fixed Version)");
        }
        public void IncrementSkippedEmails()
        {
            _skippedEmails++;
            Console.WriteLine($"🔥 DEBUG: IncrementSkippedEmails called! New count: {_skippedEmails}"); // DEBUG LINE ADDED
            UpdateStatisticsDisplay();
        }

        public void IncrementDuplicateEmails()
        {
            _duplicateEmails++;
            UpdateStatisticsDisplay();
        }

        private void CreateFloatingStatsForm()
        {
            try
            {
                // Close existing form
                if (_floatingStatsForm != null && !_floatingStatsForm.IsDisposed)
                {
                    _floatingStatsForm.Close();
                    _floatingStatsForm.Dispose();
                }

                // Create main form - NARROWER statistics form as requested
                _floatingStatsForm = new Form
                {
                    Text = "",
                    Size = new Size(450, 280), // CHANGED: Made narrower (was 600)
                    FormBorderStyle = FormBorderStyle.FixedToolWindow,
                    StartPosition = FormStartPosition.Manual,
                    BackColor = Color.FromArgb(245, 245, 245),
                    MinimizeBox = false,
                    MaximizeBox = false,
                    ShowInTaskbar = false,
                    TopMost = true
                };

                // Position relative to main form
                var parentLocation = _parentForm.Location;
                var parentSize = _parentForm.Size;
                _floatingStatsForm.Location = new Point(
                    parentLocation.X + parentSize.Width - 470, // ADJUSTED: For narrower form
                    parentLocation.Y + 50
                );

                CreateHeader();
                CreateTimePanel();
                CreateDebugConsole();
                CreateStatisticsGrid();

                _floatingStatsForm.Show(_parentForm);
                _floatingStatsForm.BringToFront();

                Console.WriteLine("✅ Clean floating stats form created (narrower)");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error creating stats form: {ex.Message}");
            }
        }

        private void CreateDebugConsole()
        {
            var debugPanel = new Panel
            {
                Location = new Point(15, 45),
                Size = new Size(280, 25), // CHANGED: Narrower debug console (was 350)
                BackColor = Color.FromArgb(240, 240, 240),
                BorderStyle = BorderStyle.FixedSingle
            };

            _floatingCurrentOperationLabel = new Label
            {
                Text = "🔄 Initializing...",
                Location = new Point(5, 3),
                Size = new Size(270, 20), // CHANGED: Narrower label (was 340)
                Font = new Font("Consolas", 8F),
                ForeColor = Color.FromArgb(64, 64, 64),
                TextAlign = ContentAlignment.MiddleLeft
            };

            debugPanel.Controls.Add(_floatingCurrentOperationLabel);
            _floatingStatsForm.Controls.Add(debugPanel);
        }

        private void CreateTimePanel()
        {
            _floatingTimeElapsedLabel = new Label
            {
                Text = "⏱️ Time: 0s",
                Location = new Point(305, 47), // CHANGED: Adjusted position for narrower form (was 375)
                Size = new Size(120, 25),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(76, 175, 80),
                TextAlign = ContentAlignment.MiddleLeft
            };

            _floatingStatsForm.Controls.Add(_floatingTimeElapsedLabel);
        }

        private void CreateStatisticsGrid()
        {
            int startY = 105;
            int rowHeight = 25;
            int colWidth = 85; // CHANGED: Narrower columns (was 100)
            int[] colX = { 15, 100, 185, 270, 355 }; // CHANGED: Adjusted for narrower form (removed 6th column)

            // Row 1: Orders and Emails
            _headerOrdersLabel = CreateGridLabel("H.Orders: 0", colX[0], startY, colWidth, rowHeight, Color.FromArgb(33, 150, 243));
            _lineOrdersLabel = CreateGridLabel("L.Orders: 0", colX[1], startY, colWidth, rowHeight, Color.FromArgb(33, 150, 243));
            _totalEmailsLabel = CreateGridLabel("T.Emails: 0", colX[2], startY, colWidth, rowHeight, Color.FromArgb(156, 39, 176));
            _skippedEmailsLabel = CreateGridLabel("S.Emails: 0", colX[3], startY, colWidth, rowHeight, Color.FromArgb(156, 39, 176));
            _duplicateEmailsLabel = CreateGridLabel("D.Emails: 0", colX[4], startY, colWidth, rowHeight, Color.FromArgb(156, 39, 176));

            // Row 2: Quotes and Errors
            startY += rowHeight + 5;
            _headerQuotesLabel = CreateGridLabel("H.Quotes: 0", colX[0], startY, colWidth, rowHeight, Color.FromArgb(255, 152, 0));
            _lineQuotesLabel = CreateGridLabel("L.Quotes: 0", colX[1], startY, colWidth, rowHeight, Color.FromArgb(255, 152, 0));
            _injectionErrorsLabel = CreateGridLabel("I.Errors: 0", colX[2], startY, colWidth, rowHeight, Color.FromArgb(244, 67, 54));
            _businessErrorsLabel = CreateGridLabel("B.Errors: 0", colX[3], startY, colWidth, rowHeight, Color.FromArgb(244, 67, 54));
            _duplicateErrorsLabel = CreateGridLabel("D.Errors: 0", colX[4], startY, colWidth, rowHeight, Color.FromArgb(244, 67, 54));

            // Row 3: Revisions
            startY += rowHeight + 5;
            _headerRevisionsLabel = CreateGridLabel("H.Revisions: 0", colX[0], startY, colWidth, rowHeight, Color.FromArgb(76, 175, 80));
            _lineRevisionsLabel = CreateGridLabel("L.Revisions: 0", colX[1], startY, colWidth, rowHeight, Color.FromArgb(76, 175, 80));
        }

        /// <summary>
        /// Update all statistics labels - INCLUDES DEBUG FOR S.EMAILS
        /// </summary>
        private void UpdateStatisticsDisplay()
        {
            try
            {
                if (_floatingStatsForm?.IsDisposed == false)
                {
                    // Update all labels
                    _headerOrdersLabel.Text = $"H.Orders: {_headerOrders}";
                    _lineOrdersLabel.Text = $"L.Orders: {_lineOrders}";
                    _totalEmailsLabel.Text = $"T.Emails: {_totalEmails}";

                    // FIXED: S.Emails counter with debug
                    _skippedEmailsLabel.Text = $"S.Emails: {_skippedEmails}";
                    Console.WriteLine($"🔥 DEBUG: S.Emails display updated to: {_skippedEmails}"); // DEBUG LINE

                    _duplicateEmailsLabel.Text = $"D.Emails: {_duplicateEmails}";
                    _headerQuotesLabel.Text = $"H.Quotes: {_headerQuotes}";
                    _lineQuotesLabel.Text = $"L.Quotes: {_lineQuotes}";
                    _injectionErrorsLabel.Text = $"I.Errors: {_injectionErrors}";
                    _businessErrorsLabel.Text = $"B.Errors: {_businessErrors}";
                    _duplicateErrorsLabel.Text = $"D.Errors: {_duplicateErrors}";
                    _headerRevisionsLabel.Text = $"H.Revisions: {_headerRevisions}";
                    _lineRevisionsLabel.Text = $"L.Revisions: {_lineRevisions}";

                    // Update time
                    var elapsed = DateTime.Now - _operationStartTime;
                    _floatingTimeElapsedLabel.Text = $"⏱️ Time: {elapsed.TotalSeconds:F0}s";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating statistics: {ex.Message}");
            }
        }
        public void IncrementDuplicateErrors()
        {
            _duplicateErrors++;
            // ADD THIS DEBUG LINE
            Console.WriteLine($"🔥 DEBUG: IncrementDuplicateErrors called! New count: {_duplicateErrors}");
            UpdateStatisticsDisplay();
        }
        public void IncrementTotalEmails()
        {
            _totalEmails++;
            UpdateStatisticsDisplay();
        }
        public void IncrementBusinessErrors(string source = "", string errorDetails = "")
        {
            _businessErrors++;

            // Optional: Internal tracking for future analysis (but unified display)
            if (source.Contains("Pre", StringComparison.OrdinalIgnoreCase) ||
                source.Contains("Validation", StringComparison.OrdinalIgnoreCase))
                _preValidationErrors++;

            if (source.Contains("MKG", StringComparison.OrdinalIgnoreCase) ||
                source.Contains("Response", StringComparison.OrdinalIgnoreCase))
                _mkgResponseErrors++;

            UpdateStatisticsDisplay();

            // Log for debugging
            LogOperation($"🚨 B.Errors++ ({_businessErrors}) - Source: {source} - {errorDetails}", LogLevel.Warning);
        }
        public void IncrementInjectionErrors()
        {
            _injectionErrors++;
            UpdateStatisticsDisplay();
        }
        #endregion

        /// <summary>
        /// Start a new session - ONLY for the very first operation
        /// </summary>
        public void StartNewSession(string sessionName, int totalSteps)
        {
            _operationStartTime = DateTime.Now;
            _currentOperation = sessionName;
            _totalSteps = totalSteps;
            _currentStep = 0;
            _isProcessingComplete = false;
            _preventAutoHide = true; // 🎯 PREVENT auto-hide for multi-phase workflows

            // Reset counters only for a brand new session
            ResetAllCounters();

            // Create the floating form (only once)
            CreateFloatingStatsForm();

            UpdateMainProgress(0, totalSteps, $"Starting {sessionName}...");
            UpdateDebugConsole($"🚀 Starting {sessionName}...");
            LogOperation($"🚀 New session started: {sessionName}", LogLevel.Info);

            _updateTimer.Start();
        }

        /// <summary>
        /// Update operation phase - does NOT recreate form or reset counters
        /// Use this for switching between phases (email import → MKG injection)
        /// </summary>
        public void UpdateOperationPhase(string phaseName, int currentStep, int totalSteps)
        {
            _currentStep = currentStep;
            _totalSteps = totalSteps;

            // 🎯 FIXED: Ensure form is visible for new phase
            EnsureStatsFormVisible();

            UpdateMainProgress(currentStep, totalSteps, phaseName);
            UpdateDebugConsole($"🔄 {phaseName}");
            LogOperation($"📊 Phase update: {phaseName}", LogLevel.Info);
        }

        /// <summary>
        /// Legacy StartOperation - now properly handles phases
        /// </summary>
        public void StartOperation(string operationName, int totalSteps)
        {
            // If form doesn't exist, create it (first time only)
            if (_floatingStatsForm == null || _floatingStatsForm.IsDisposed)
            {
                StartNewSession(operationName, totalSteps);
            }
            else
            {
                // Form exists, just update the phase
                UpdateOperationPhase(operationName, 0, totalSteps);
            }
        }

        /// <summary>
        /// Update progress during operation - FIXED VERSION
        /// </summary>
        public void UpdateProgress(int current, string status, string subOperation = null)
        {
            _currentStep = current;
            _lastOperation = status ?? "Processing...";

            // 🎯 FIXED: Only update debug console if subOperation is explicitly provided
            if (!string.IsNullOrEmpty(subOperation))
            {
                UpdateDebugConsole(subOperation);
            }

            UpdateMainProgress(current, _totalSteps, status);
            UpdateStatisticsDisplay(); // This now WON'T overwrite debug console

            LogOperation($"📊 Progress: {current}/{_totalSteps} - {status}", LogLevel.Info);
        }

        /// <summary>
        /// Update progress with total - FIXED VERSION
        /// </summary>
        public void UpdateProgress(int current, int total, string status)
        {
            _currentStep = current;
            _totalSteps = total;
            _lastOperation = status ?? "Processing...";

            UpdateMainProgress(current, total, status);
            UpdateStatisticsDisplay(); // This now WON'T overwrite debug console

            LogOperation($"📊 Progress: {current}/{total} - {status}", LogLevel.Info);
        }

        /// <summary>
        /// Complete operation - FIXED to NOT auto-hide during multi-phase workflows
        /// </summary>
        public void CompleteOperation(string message = null)
        {
            var elapsed = DateTime.Now - _operationStartTime;
            var completionMessage = message ?? "Operation completed successfully";

            UpdateMainProgress(_totalSteps, _totalSteps, $"✅ {completionMessage}");
            LogOperation($"✅ Completed in {elapsed.TotalSeconds:F1}s - {completionMessage}", LogLevel.Success);

            // 🎯 FIXED: Don't stop timer or hide form if preventing auto-hide
            if (!_preventAutoHide)
            {
                _updateTimer.Stop();
                _isProcessingComplete = true;
            }

            // Update debug console for completion
            UpdateDebugConsole($"✅ {completionMessage}");

            // Change color to show completion
            if (_floatingCurrentOperationLabel != null)
            {
                _floatingCurrentOperationLabel.ForeColor = Color.DarkGreen;
            }

            // 🎯 FIXED: Only auto-hide if NOT preventing
            if (!_preventAutoHide)
            {
                Task.Delay(3000).ContinueWith(_ => HideFloatingStatsForm());
            }
        }

        /// <summary>
        /// NEW: Complete entire workflow - call this when EVERYTHING is done
        /// </summary>
        public void CompleteEntireWorkflow(string finalMessage = null)
        {
            var elapsed = DateTime.Now - _operationStartTime;
            var completionMessage = finalMessage ?? "All operations completed successfully";

            _preventAutoHide = false; // Allow hiding now
            _isProcessingComplete = true;
            _updateTimer.Stop();

            UpdateMainProgress(_totalSteps, _totalSteps, $"🎉 {completionMessage}");
            UpdateDebugConsole($"🎉 {completionMessage}");
            LogOperation($"🎉 Complete workflow finished in {elapsed.TotalSeconds:F1}s", LogLevel.Success);

            // Change color to show final completion
            if (_floatingCurrentOperationLabel != null)
            {
                _floatingCurrentOperationLabel.ForeColor = Color.DarkGreen;
            }

            // Auto-hide after 5 seconds
            Task.Delay(5000).ContinueWith(_ => HideFloatingStatsForm());
        }

        /// <summary>
        /// Enhanced fail operation
        /// </summary>
        public void FailOperation(string errorMessage)
        {
            _lastOperation = $"FAILED: {errorMessage}";

            UpdateDebugConsole($"❌ FAILED: {errorMessage}");

            // Change color to show error
            if (_floatingCurrentOperationLabel != null)
            {
                _floatingCurrentOperationLabel.ForeColor = Color.DarkRed;
            }

            _statusLabel.Text = $"Failed: {errorMessage}";
            LogOperation($"❌ Operation failed: {errorMessage}", LogLevel.Error);
        }

        #region Statistics Increment Methods (Real-time Updates)

        public void IncrementEmailsProcessed()
        {
            _emailsProcessed++;
            UpdateStatisticsDisplay();
        }

        public void IncrementOrdersFound(int count = 1)
        {
            _ordersFound += count;
            UpdateStatisticsDisplay();
        }

        public void IncrementQuotesFound(int count = 1)
        {
            _quotesFound += count;
            UpdateStatisticsDisplay();
        }

        public void IncrementRevisionsFound(int count = 1)
        {
            _revisionsFound += count;
            UpdateStatisticsDisplay();
        }

        public void IncrementHeaderOrders(int count = 1)
        {
            _headerOrders += count;
            UpdateStatisticsDisplay();
        }

        public void IncrementLineOrders(int count = 1)
        {
            _lineOrders += count;
            UpdateStatisticsDisplay();
        }

        public void IncrementHeaderQuotes(int count = 1)
        {
            _headerQuotes += count;
            UpdateStatisticsDisplay();
        }

        public void IncrementLineQuotes(int count = 1)
        {
            _lineQuotes += count;
            UpdateStatisticsDisplay();
        }

        public void IncrementHeaderRevisions(int count = 1)
        {
            _headerRevisions += count;
            UpdateStatisticsDisplay();
        }

        public void IncrementLineRevisions(int count = 1)
        {
            _lineRevisions += count;
            UpdateStatisticsDisplay();
        }

        #endregion

        #region Statistics Update Methods (CLEAN API)

        public void UpdateEmailStats(int total, int skipped = 0, int duplicates = 0)
        {
            _totalEmails = total;
            _skippedEmails = skipped;
            _duplicateEmails = duplicates;
            UpdateStatisticsDisplay();
        }

        public void UpdateOrderStats(int headers, int lines)
        {
            _headerOrders = headers;
            _lineOrders = lines;
            UpdateStatisticsDisplay();
        }

        public void UpdateQuoteStats(int headers, int lines)
        {
            _headerQuotes = headers;
            _lineQuotes = lines;
            UpdateStatisticsDisplay();
        }

        public void UpdateRevisionStats(int headers, int lines)
        {
            _headerRevisions = headers;
            _lineRevisions = lines;
            UpdateStatisticsDisplay();
        }

        public void UpdateErrorStats(int injection = 0, int business = 0, int duplicates = 0)
        {
            _injectionErrors = injection;
            _businessErrors = business;
            _duplicateErrors = duplicates;
            UpdateStatisticsDisplay();
        }

        #endregion

        #region Debug Console Management - FIXED

        /// <summary>
        /// Enhanced UpdateDebugConsole - THE ONLY method that updates debug console
        /// </summary>
        public void UpdateDebugConsole(string message)
        {
            try
            {
                if (_floatingCurrentOperationLabel != null)
                {
                    if (_floatingStatsForm.InvokeRequired)
                    {
                        _floatingStatsForm.Invoke((Action)(() => _floatingCurrentOperationLabel.Text = message));
                    }
                    else
                    {
                        _floatingCurrentOperationLabel.Text = message;
                    }

                    // Keep _currentOperation in sync for logging
                    _currentOperation = message.Replace("🔄 ", "").Replace("📧 ", "").Replace("🚀 ", "").Replace("✅ ", "");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating debug console: {ex.Message}");
            }
        }

        #endregion

        #region Floating Statistics Form

        /// <summary>
        /// Create the clean floating stats window
        /// </summary>
       
        private void CreateHeader()
        {
            var headerPanel = new Panel
            {
                Size = new Size(_floatingStatsForm.ClientSize.Width, 32),
                Location = new Point(0, 0),
                BackColor = Color.FromArgb(76, 175, 80),
                Dock = DockStyle.Top
            };

            var headerLabel = new Label
            {
                Text = "📊 LIVE STATISTICS - PROCESSING ACTIVE",
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                ForeColor = Color.White,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            };

            headerPanel.Controls.Add(headerLabel);
            _floatingStatsForm.Controls.Add(headerPanel);
        }

      
        private Label CreateGridLabel(string text, int x, int y, int width, int height, Color foreColor)
        {
            var label = new Label
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(width, height),
                Font = new Font("Segoe UI", 8F, FontStyle.Regular),
                ForeColor = foreColor,
                BackColor = Color.Transparent,
                TextAlign = ContentAlignment.MiddleLeft
            };

            _floatingStatsForm.Controls.Add(label);
            return label;
        }

        #endregion

        #region Statistics Display Update - FIXED VERSION

        /// <summary>
        /// Update all statistics labels - FIXED to NOT overwrite debug console
        /// </summary>
       
        #endregion

        #region Utility Methods

        private void InitializeUpdateTimer()
        {
            _updateTimer = new Timer();
            _updateTimer.Interval = 500;
            _updateTimer.Tick += (s, e) => UpdateStatisticsDisplay();
        }

        private void UpdateMainProgress(int current, int total, string status)
        {
            try
            {
                if (_parentForm.InvokeRequired)
                {
                    _parentForm.Invoke((Action)(() => UpdateMainProgress(current, total, status)));
                    return;
                }

                if (total > 0)
                {
                    var percentage = Math.Min(100, (current * 100) / total);
                    _progressBar.Value = percentage;
                    _progressLabel.Text = $"{current}/{total}";
                }

                _statusLabel.Text = status;
                Application.DoEvents();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating progress: {ex.Message}");
            }
        }

        /// <summary>
        /// Ensure statistics form is visible and functional
        /// </summary>
        private void EnsureStatsFormVisible()
        {
            try
            {
                if (_floatingStatsForm == null || _floatingStatsForm.IsDisposed)
                {
                    CreateFloatingStatsForm();
                    if (!_updateTimer.Enabled)
                    {
                        _updateTimer.Start();
                    }
                }
                else if (!_floatingStatsForm.Visible)
                {
                    _floatingStatsForm.Show(_parentForm);
                    _floatingStatsForm.BringToFront();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error ensuring stats form visible: {ex.Message}");
            }
        }

        private void HideFloatingStatsForm()
        {
            try
            {
                if (_floatingStatsForm?.IsDisposed == false && _floatingStatsForm.Visible)
                {
                    if (_floatingStatsForm.InvokeRequired)
                    {
                        _floatingStatsForm.Invoke((Action)HideFloatingStatsForm);
                        return;
                    }

                    // Only hide if processing is actually complete
                    if (_isProcessingComplete && !_preventAutoHide)
                    {
                        _floatingStatsForm.Hide();
                        LogOperation("👁️ Statistics form hidden after completion", LogLevel.Info);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error hiding stats form: {ex.Message}");
            }
        }

        private void ResetAllCounters()
        {
            _headerOrders = _lineOrders = _totalEmails = _skippedEmails = _duplicateEmails = 0;
            _headerQuotes = _lineQuotes = _injectionErrors = _businessErrors = _duplicateErrors = 0;
            _headerRevisions = _lineRevisions = 0;
            _emailsProcessed = _ordersFound = _quotesFound = _revisionsFound = 0;
            _currentOperation = "Starting...";
        }

        private void LogOperation(string message, LogLevel level)
        {
            try
            {
                LogEntryAdded?.Invoke(this, new LogEntryEventArgs { Message = message, Level = level });
                Console.WriteLine($"[{level}] {message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error logging: {ex.Message}");
            }
        }

        #endregion

        #region Public Properties

        public TimeSpan ElapsedTime => DateTime.Now - _operationStartTime;
        public string CurrentOperation => _currentOperation;
        public int TotalEmails => _totalEmails;
        public int HeaderOrders => _headerOrders;
        public int LineOrders => _lineOrders;
        public int HeaderQuotes => _headerQuotes;
        public int LineQuotes => _lineQuotes;

        public int _preValidationErrors { get; private set; }
        public int _mkgResponseErrors { get; private set; }

        #endregion

        #region Disposal

        public void Dispose()
        {
            try
            {
                _updateTimer?.Stop();
                _updateTimer?.Dispose();

                if (_floatingStatsForm != null && !_floatingStatsForm.IsDisposed)
                {
                    _floatingStatsForm.Close();
                    _floatingStatsForm.Dispose();
                }

                _tooltip?.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error disposing: {ex.Message}");
            }
        }

        #endregion
    }

    /// <summary>
    /// Enhanced Progress Callback for seamless integration
    /// </summary>
    public class EnhancedProgressCallback : IProgress<(int current, int total, string status)>
    {
        private readonly EnhancedProgressManager _progressManager;

        public EnhancedProgressCallback(EnhancedProgressManager progressManager)
        {
            _progressManager = progressManager;
        }

        public void Report((int current, int total, string status) value)
        {
            _progressManager.UpdateProgress(value.current, value.total, value.status);
        }
    }
}