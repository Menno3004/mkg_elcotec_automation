using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using Mkg_Elcotec_Automation.Services;
using Mkg_Elcotec_Automation.Models;
using Mkg_Elcotec_Automation.Models.EmailModels;
using Mkg_Elcotec_Automation.Controllers;

namespace Mkg_Elcotec_Automation.Forms
{
    
    public partial class Elcotec : Form
    {
        #region Private Fields
        private EnhancedProgressManager _enhancedProgress;
        private bool _isProcessing = false;
        private DateTime _processingStartTime;

        // Bottom ToolStrip components
        private ToolStrip bottomToolStrip;
        private ToolStripProgressBar toolStripProgressBar;
        private ToolStripStatusLabel toolStripStatusLabel;
        private ToolStripLabel toolStripProgressLabel;

        // Form components
        private ListBox lstEmailImportResults;
        private ListBox lstMkgResults;
        private ListBox lstFailedInjections;
        private Label lblMkgStatus;
        private Label lblFailedStatus;

        private MkgOrderInjectionSummary _lastOrderSummary;
        private MkgQuoteInjectionSummary _lastQuoteSummary;
        private MkgRevisionInjectionSummary _lastRevisionSummary;

        // 🎯 ADD: Run History Fields (add after your existing private fields)
        private ComboBox cmbRunHistory;
        private RunHistoryManager _runHistoryManager;
        private Guid _currentRunId;
        private bool _isLoadingHistoricalRun = false;
        private AutomationRunData _currentRunData;
        private bool _hasMkgDuplicatesDetected = false;
        private int _totalDuplicatesInCurrentRun = 0;
        private bool _hasInjectionFailures = false;
        private List<string> _savedIncrementalOutput = new List<string>();
        #endregion

        #region Constructor and Initialization
        public Elcotec()
        {
            InitializeComponent();
            tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
            tabControl.DrawItem += TabControl_DrawItem;
            SetupEnhancedDuplicationEventHandlers();
            SetupMkgMenuStrip();
            SetupBottomToolStrip();
            InitializeRunHistory();
            InitializeAllTabControls();
            InitializeEnhancedProgress();
            SetupEnhancedDuplicationEventHandlers();
            Console.WriteLine("✅ Elcotec constructor completed with event wiring");
        }
        // Add these 3 simple methods to Elcotec.cs
        private void DisplayOrderFailures(MkgOrderInjectionSummary summary)
        {
            var realFailures = summary.OrderResults.Where(r => !r.Success && !r.HttpStatusCode.Contains("DUPLICATE")).ToList();
            if (!realFailures.Any()) return;

            lstFailedInjections.Items.Add($"📦 ORDER FAILURES ({realFailures.Count}):");
            foreach (var failure in realFailures.Take(10)) // Limit display
            {
                lstFailedInjections.Items.Add($"   • {failure.ArtiCode}: {failure.ErrorMessage}");
            }
            if (realFailures.Count > 10)
                lstFailedInjections.Items.Add($"   ... and {realFailures.Count - 10} more order failures");
            lstFailedInjections.Items.Add("");
        }

        private void DisplayQuoteFailures(MkgQuoteInjectionSummary summary)
        {
            var realFailures = summary.QuoteResults.Where(r => !r.Success && !r.HttpStatusCode.Contains("DUPLICATE")).ToList();
            if (!realFailures.Any()) return;

            lstFailedInjections.Items.Add($"💰 QUOTE FAILURES ({realFailures.Count}):");
            foreach (var failure in realFailures.Take(10))
            {
                lstFailedInjections.Items.Add($"   • {failure.ArtiCode}: {failure.ErrorMessage}");
            }
            if (realFailures.Count > 10)
                lstFailedInjections.Items.Add($"   ... and {realFailures.Count - 10} more quote failures");
            lstFailedInjections.Items.Add("");
        }

        private void DisplayRevisionFailures(MkgRevisionInjectionSummary summary)
        {
            var realFailures = summary.RevisionResults.Where(r => !r.Success && !r.HttpStatusCode.Contains("DUPLICATE")).ToList();
            if (!realFailures.Any()) return;

            lstFailedInjections.Items.Add($"🔄 REVISION FAILURES ({realFailures.Count}):");
            foreach (var failure in realFailures.Take(10))
            {
                lstFailedInjections.Items.Add($"   • {failure.ArtiCode}: {failure.ErrorMessage}");
            }
            if (realFailures.Count > 10)
                lstFailedInjections.Items.Add($"   ... and {realFailures.Count - 10} more revision failures");
            lstFailedInjections.Items.Add("");
        }
        private void ShowInjectionErrors(MkgOrderInjectionSummary orderSummary, MkgQuoteInjectionSummary quoteSummary, MkgRevisionInjectionSummary revisionSummary)
        {
            var count = _enhancedProgress?.GetInjectionErrors() ?? 0;
            lstFailedInjections.Items.Add($"🔧 I:Errors: {count} - Technical injection failures");

            if (count > 0)
            {
                if (orderSummary?.OrderResults != null)
                {
                    var failures = orderSummary.OrderResults.Where(r => !r.Success && r.HttpStatusCode != "MKG_DUPLICATE_SKIPPED").Take(3);
                    foreach (var f in failures)
                        lstFailedInjections.Items.Add($"   • {f.ArtiCode}: {f.ErrorMessage}");
                }
            }
            lstFailedInjections.Items.Add("");
        }
        // 🔧 FIXED: Event handlers for duplication service - Do NOT set injection failures for duplicates
        private void SetupDuplicationServiceEvents()
        {
            try
            {
                // 🚨 CRITICAL BUSINESS ERRORS - Financial Protection - RED tab
                DuplicationService.OnCriticalBusinessErrors += (count) =>
                {
                    try
                    {
                        Console.WriteLine($"🚨 CRITICAL BUSINESS ERRORS: {count} detected - immediate action required");
                        _hasInjectionFailures = true; // ✅ CORRECT: Force RED tab for business errors
                        UpdateTabStatus();

                        // Update current run data
                        if (_currentRunData != null)
                        {
                            _currentRunData.HasInjectionFailures = true;
                            _currentRunData.TotalFailuresAtCompletion += count;
                        }

                        // Log critical errors in Failed Injections tab
                        LogResult($"🚨 CRITICAL BUSINESS ERRORS: {count} detected - immediate action required");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error handling critical business errors: {ex.Message}");
                    }
                };

                // 💰 PRICE DISCREPANCY ALERTS - RED tab  
                DuplicationService.OnPriceDiscrepancyDetected += (alert) =>
                {
                    try
                    {
                        Console.WriteLine($"💰 PRICE DISCREPANCY: {alert.ArtiCode} - {alert.DiscrepancyPercentage:F1}% difference");
                        _hasInjectionFailures = true; // ✅ CORRECT: Force RED tab for financial risks
                        UpdateTabStatus();

                        // Update current run data
                        if (_currentRunData != null)
                        {
                            _currentRunData.HasInjectionFailures = true;
                            _currentRunData.TotalFailuresAtCompletion++;
                        }

                        // Log price discrepancy in Failed Injections tab
                        LogResult($"💰 PRICE DISCREPANCY: {alert.ArtiCode} - {alert.DiscrepancyPercentage:F1}% difference ({alert.FinancialRisk} risk)");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error handling price discrepancy: {ex.Message}");
                    }
                };

                // 🔄 MANAGED MKG DUPLICATES - YELLOW tab ONLY (NOT RED)
                DuplicationService.OnManagedMkgDuplicates += (count) =>
                {
                    try
                    {
                        Console.WriteLine($"🔄 MANAGED DUPLICATES DETECTED: {count} (NOT an error)");

                        // 🔧 FIXED: Do NOT set _hasInjectionFailures for duplicates
                        // _hasInjectionFailures = true; // ❌ WRONG - removed this line

                        _hasMkgDuplicatesDetected = true; // ✅ CORRECT: Only set duplicate flag
                        _totalDuplicatesInCurrentRun += count;
                        UpdateTabStatus();

                        // Update current run data - duplicates are NOT failures
                        if (_currentRunData != null)
                        {
                            _currentRunData.HasMkgDuplicatesDetected = true;
                            _currentRunData.TotalDuplicatesDetected += count;
                            _currentRunData.LastDuplicateDetectionUpdate = DateTime.Now;

                            // 🔧 FIXED: Do NOT increment failure count for duplicates
                            // _currentRunData.HasInjectionFailures = true; // ❌ WRONG - removed
                            // _currentRunData.TotalFailuresAtCompletion++; // ❌ WRONG - removed
                        }

                        // Log managed duplicates
                        LogResult($"🔄 MANAGED DUPLICATES: {count} items handled automatically");
                        Console.WriteLine($"✅ Duplicates processed: Total={_totalDuplicatesInCurrentRun}, Tab should be YELLOW (not red)");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error handling managed duplicates: {ex.Message}");
                    }
                };

                Console.WriteLine("✅ Duplication service events configured - duplicates will NOT cause red tab");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error setting up duplication service events: {ex.Message}");
            }
        }
        // In Elcotec.cs - Remove the calls to UpdateFailedInjectionsTabHeaderColor during email import

        // REMOVE this line from DisplayBasicHistoricalInfo method:
        // UpdateFailedInjectionsTabHeaderColor(runInfo);

        // REMOVE this line from DisplayHistoricalRunData method:  
        // UpdateFailedInjectionsTabHeaderColor(historicalData);

        // Also remove from UpdateStatusLabelsForDetailedHistory - replace UpdateTabColorsForHistoricalData call
        // with conditional logic that only updates after MKG injection

        // The UpdateFailedInjectionsTabHeaderColor method should ONLY be called:
        // 1. After MKG injection is complete
        // 2. When loading completed historical runs (not during import)

        private void DisplayBasicHistoricalInfo(RunHistoryItem runInfo)
        {
            try
            {
                Console.WriteLine($"🔄 Displaying basic historical info for run: {runInfo.RunId}");

                // Clear current displays
                lstEmailImportResults?.Items.Clear();
                lstMkgResults?.Items.Clear();
                lstFailedInjections?.Items.Clear();

                // Display basic info only
                if (runInfo.Settings != null)
                {
                    lstEmailImportResults.Items.Add($"📧 Run ID: {runInfo.RunId}");
                    lstEmailImportResults.Items.Add($"📅 Date: {runInfo.StartTime:yyyy-MM-dd HH:mm:ss}");
                    lstEmailImportResults.Items.Add($"⏱️ Duration: {runInfo.Duration}");
                    lstEmailImportResults.Items.Add($"📊 Status: {runInfo.Status}");
                }

                // Update status labels
                UpdateStatusLabelsForBasicHistory(runInfo);

                // 🔧 REMOVED: Don't update tab color during email import phase
                // Only update tab color if this is a completed run (status = "Completed" or "Failed")
                if (runInfo.Status == "Completed" || runInfo.Status == "Failed")
                {
                    UpdateFailedInjectionsTabHeaderColor(runInfo);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error displaying basic historical info: {ex.Message}");
            }
        }

        private void DisplayHistoricalRunData(AutomationRunData historicalData)
        {
            try
            {
                Console.WriteLine("🔄 Updating UI with detailed historical data...");

                // Clear current displays
                lstEmailImportResults?.Items.Clear();
                lstMkgResults?.Items.Clear();
                lstFailedInjections?.Items.Clear();

                // 1. Email Import Results Tab - using EXISTING method
                DisplayHistoricalEmailResults(historicalData);

                // 2. MKG Results Tab - using EXISTING method
                DisplayHistoricalMkgResults(historicalData);

                // 3. Failed Injections Tab - using EXISTING method
                DisplayHistoricalFailedInjections(historicalData);

                // 4. Update status labels - using EXISTING method
                UpdateStatusLabelsForDetailedHistory(historicalData);

                // 5. 🔧 REMOVED: Don't update tab color during email import phase
                // Only update tab color if this historical data represents a completed run
                // Check if the run is completed by looking at RunInfo status
                if (historicalData.RunInfo != null &&
                    (historicalData.RunInfo.Status == "Completed" || historicalData.RunInfo.Status == "Failed"))
                {
                    UpdateFailedInjectionsTabHeaderColor(historicalData);
                }

                Console.WriteLine("✅ UI updated with detailed historical data");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error displaying detailed historical data: {ex.Message}");
            }
        }

        // The issue is in UpdateFailedInjectionsTab method
        // The calculation for totalRealFailures is wrong, resulting in negative numbers
        // This causes _hasInjectionFailures to be set incorrectly

        // Fix the calculation in UpdateFailedInjectionsTab method:

        public void UpdateFailedInjectionsTab(MkgOrderInjectionSummary orderSummary = null,
                                MkgQuoteInjectionSummary quoteSummary = null,
                                MkgRevisionInjectionSummary revisionSummary = null)
        {
            try
            {
                if (lstFailedInjections == null) return;

                lstFailedInjections.Items.Clear();

                // Calculate totals correctly
                var totalOrderResults = orderSummary?.OrderResults?.Count ?? 0;
                var totalQuoteResults = quoteSummary?.QuoteResults?.Count ?? 0;
                var totalRevisionResults = revisionSummary?.RevisionResults?.Count ?? 0;

                // 🔧 CRITICAL FIX: Count duplicates and real failures correctly
                var totalMkgDuplicates = 0;
                var totalRealFailures = 0;

                // Count order failures correctly
                if (orderSummary?.OrderResults != null)
                {
                    foreach (var result in orderSummary.OrderResults)
                    {
                        if (!result.Success)
                        {
                            if (result.HttpStatusCode == "MKG_DUPLICATE_SKIPPED" ||
                                result.ErrorMessage?.Contains("DUPLICATE") == true ||
                                result.ErrorMessage?.Contains("already exists") == true)
                            {
                                totalMkgDuplicates++;
                            }
                            else
                            {
                                totalRealFailures++;
                            }
                        }
                    }
                }

                // Count quote failures correctly  
                if (quoteSummary?.QuoteResults != null)
                {
                    foreach (var result in quoteSummary.QuoteResults)
                    {
                        if (!result.Success)
                        {
                            if (result.HttpStatusCode == "MKG_DUPLICATE_SKIPPED" ||
                                result.ErrorMessage?.Contains("DUPLICATE") == true ||
                                result.ErrorMessage?.Contains("already exists") == true)
                            {
                                totalMkgDuplicates++;
                            }
                            else
                            {
                                totalRealFailures++;
                            }
                        }
                    }
                }

                // Count revision failures correctly
                if (revisionSummary?.RevisionResults != null)
                {
                    foreach (var result in revisionSummary.RevisionResults)
                    {
                        if (!result.Success)
                        {
                            if (result.HttpStatusCode == "MKG_DUPLICATE_SKIPPED" ||
                                result.ErrorMessage?.Contains("DUPLICATE") == true ||
                                result.ErrorMessage?.Contains("already exists") == true)
                            {
                                totalMkgDuplicates++;
                            }
                            else
                            {
                                totalRealFailures++;
                            }
                        }
                    }
                }

                // 🔧 CRITICAL FIX: Set flags correctly based on counts
                _hasInjectionFailures = (totalRealFailures > 0);  // Only true for REAL failures
                _hasMkgDuplicatesDetected = (totalMkgDuplicates > 0);  // True for duplicates
                _totalDuplicatesInCurrentRun = totalMkgDuplicates;

                Console.WriteLine($"🔧 FIXED FLAGS: _hasInjectionFailures={_hasInjectionFailures}, _hasMkgDuplicatesDetected={_hasMkgDuplicatesDetected}");
                Console.WriteLine($"🔧 COUNTS: Real failures={totalRealFailures}, Duplicates={totalMkgDuplicates}");

                // 🔧 FIX: Update current run data correctly
                if (_currentRunData != null)
                {
                    _currentRunData.HasInjectionFailures = (totalRealFailures > 0);
                    _currentRunData.TotalFailuresAtCompletion = totalRealFailures;
                    _currentRunData.HasMkgDuplicatesDetected = (totalMkgDuplicates > 0);
                    _currentRunData.TotalDuplicatesDetected = totalMkgDuplicates;
                }

                // Update status label with correct colors
                if (lblFailedStatus != null)
                {
                    if (totalRealFailures == 0 && totalMkgDuplicates == 0)
                    {
                        lblFailedStatus.Text = "✅ No issues detected - all injections successful!";
                        lblFailedStatus.ForeColor = Color.Green;
                    }
                    else if (totalRealFailures > 0)
                    {
                        lblFailedStatus.Text = $"❌ {totalRealFailures} real failures, {totalMkgDuplicates} duplicates detected";
                        lblFailedStatus.ForeColor = Color.Red;
                    }
                    else // Only MKG duplicates, no real failures
                    {
                        lblFailedStatus.Text = $"🔄 {totalMkgDuplicates} duplicates detected (handled automatically)";
                        lblFailedStatus.ForeColor = Color.DarkOrange;
                    }
                }

                // 🔧 FIX: Update tab status with correct values
                UpdateTabStatus();

                // Rest of the method continues as before...
                // Header
                lstFailedInjections.Items.Add("🔍 === DETAILED INJECTION FAILURE ANALYSIS ===");
                lstFailedInjections.Items.Add($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

                if (totalRealFailures == 0 && totalMkgDuplicates == 0)
                {
                    lstFailedInjections.Items.Add("✅ NO INJECTION FAILURES!");
                    lstFailedInjections.Items.Add("🎉 All items processed successfully");
                }
                else if (totalRealFailures == 0)
                {
                    lstFailedInjections.Items.Add("✅ NO REAL FAILURES DETECTED");
                    lstFailedInjections.Items.Add($"🔄 {totalMkgDuplicates} duplicates were automatically handled");
                    lstFailedInjections.Items.Add("💡 Tab should be YELLOW/ORANGE - duplicates are not errors");
                }
                else
                {
                    lstFailedInjections.Items.Add($"❌ {totalRealFailures} REAL FAILURES detected");
                    lstFailedInjections.Items.Add($"🔄 {totalMkgDuplicates} duplicates were automatically handled");
                    lstFailedInjections.Items.Add("💡 Tab should be RED - real failures require attention");
                }

                lstFailedInjections.Items.Add("");

                // Continue with existing display logic...
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating failed injections tab: {ex.Message}");
            }
        }

        // Modify UpdateTabStatus to check this flag:
        public void UpdateTabStatus()
        {
            try
            {
                var currentStatus = GetCurrentTabStatus();

                // 🔧 CRITICAL: Force redraw ONLY if we have a tab control
                if (tabControl != null)
                {
                    tabControl.Invalidate();
                    tabControl.Update();
                }

                var statusText = currentStatus switch
                {
                    TabStatus.Red => "RED (critical failures or price discrepancies)",
                    TabStatus.Yellow => "YELLOW (managed duplicates detected)",
                    TabStatus.Green => "GREEN (no issues)",
                    _ => "UNKNOWN"
                };

                Console.WriteLine($"🎨 Tab status updated to: {statusText}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating tab status: {ex.Message}");
            }
        }

        // Modify the DuplicationService event handlers to check the flag:
        // 🔧 FIXED: Proper status priority logic
        private TabStatus GetCurrentTabStatus()
        {
            // RED takes highest priority - ONLY for actual injection errors, critical business errors, or price discrepancies
            if (_hasInjectionFailures)
            {
                Console.WriteLine($"🔴 Tab status: RED - injection failures detected");
                return TabStatus.Red;
            }

            // YELLOW - duplicates detected but no actual failures
            if (_hasMkgDuplicatesDetected && _totalDuplicatesInCurrentRun > 0)
            {
                Console.WriteLine($"🟡 Tab status: YELLOW - {_totalDuplicatesInCurrentRun} duplicates detected (not errors)");
                return TabStatus.Yellow;
            }

            // GREEN - no errors, no duplicates
            Console.WriteLine($"🟢 Tab status: GREEN - no issues detected");
            return TabStatus.Green;
        }

        // 🔧 FIXED: Enhanced completion handler that properly tracks actual failures vs duplicates
        private void CompleteRunWithProperStatus()
        {
            try
            {
                if (_currentRunId != Guid.Empty)
                {
                    // 🔧 Count ONLY real failures, not duplicates
                    var realFailureCount = 0;

                    // Count actual injection errors from your results (simple string check since ErrorResults is List<object>)
                    if (_currentRunData?.ErrorResults != null)
                    {
                        realFailureCount = _currentRunData.ErrorResults
                            .Count(e => {
                                var errorText = e?.ToString() ?? "";
                                return !errorText.Contains("DUPLICATE") &&
                                       !errorText.Contains("already exists") &&
                                       !errorText.Contains("MKG_DUPLICATE_SKIPPED");
                            });
                    }

                    // Update run completion with accurate data
                    _currentRunData.TotalFailuresAtCompletion = realFailureCount;
                    _currentRunData.HasInjectionFailures = realFailureCount > 0;

                    // Complete the run with proper status
                    var finalStatus = realFailureCount > 0 ? "Failed" : "Completed";
                    _runHistoryManager.CompleteRun(_currentRunId, finalStatus);

                    Console.WriteLine($"✅ Run completed with status: {finalStatus}");
                    Console.WriteLine($"   📊 Real failures: {realFailureCount}");
                    Console.WriteLine($"   📊 Managed duplicates: {_totalDuplicatesInCurrentRun}");
                    Console.WriteLine($"   📊 Final tab should be: {(realFailureCount > 0 ? "RED" : (_totalDuplicatesInCurrentRun > 0 ? "YELLOW" : "GREEN"))}");

                    RefreshRunHistory();

                    // Force final tab update
                    UpdateTabStatus();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error completing run: {ex.Message}");
            }
        }
      
       
        
        private void ShowBusinessErrors(MkgOrderInjectionSummary orderSummary, MkgQuoteInjectionSummary quoteSummary, MkgRevisionInjectionSummary revisionSummary)
        {
            var count = _enhancedProgress?.GetBusinessErrors() ?? 0;
            lstFailedInjections.Items.Add($"🚨 B:Errors: {count} - Business rule violations");

            if (count > 0)
            {
                if (orderSummary?.OrderResults != null)
                {
                    var failures = orderSummary.OrderResults.Where(r => !r.Success && r.HttpStatusCode == "BUSINESS_ERROR").Take(3);
                    foreach (var f in failures)
                        lstFailedInjections.Items.Add($"   • {f.ArtiCode}: {f.ErrorMessage}");
                }
            }
            lstFailedInjections.Items.Add("");
        }

        private void ShowDuplicateErrors(MkgOrderInjectionSummary orderSummary, MkgQuoteInjectionSummary quoteSummary, MkgRevisionInjectionSummary revisionSummary)
        {
            var count = _enhancedProgress?.GetDuplicateErrors() ?? 0;
            lstFailedInjections.Items.Add($"🔄 D:Errors: {count} - Duplicate items skipped");

            if (count > 0)
            {
                if (orderSummary?.OrderResults != null)
                {
                    var duplicates = orderSummary.OrderResults.Where(r => r.HttpStatusCode == "MKG_DUPLICATE_SKIPPED").Take(3);
                    foreach (var d in duplicates)
                        lstFailedInjections.Items.Add($"   • {d.ArtiCode}: Already exists in MKG");
                }
            }
            lstFailedInjections.Items.Add("");
        }

        
        // ===== FIX TAB BUTTONS - 3 IDENTICAL BUTTONS ON EVERY TAB =====
        // ===== RESTORE EXISTING WORKING TAB BUTTON CODE WITH 3 BUTTONS PER TAB =====
        // ===== RESTORE EXISTING WORKING TAB BUTTON CODE WITH 3 BUTTONS PER TAB =====

        private void SetupEmailImportTab()
        {
            tabEmailImport.Controls.Clear();

            var lblEmailStatus = new Label();
            lblEmailStatus.Location = new Point(10, 10);
            lblEmailStatus.Size = new Size(600, 23);
            lblEmailStatus.Text = "Ready to import emails...";
            lblEmailStatus.Font = new Font("Arial", 11F, FontStyle.Bold);
            lblEmailStatus.ForeColor = Color.DarkBlue;
            lblEmailStatus.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            tabEmailImport.Controls.Add(lblEmailStatus);

            var buttonPanel = new Panel();
            buttonPanel.Location = new Point(10, 35);
            buttonPanel.Size = new Size(tabControl.Width - 40, 30);
            buttonPanel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            tabEmailImport.Controls.Add(buttonPanel);

            var btnCopyEmailResults = new Button();
            btnCopyEmailResults.Text = "Copy Tab";
            btnCopyEmailResults.Location = new Point(0, 0);
            btnCopyEmailResults.Size = new Size(80, 25);
            btnCopyEmailResults.Click += (s, e) => CopyListBoxToClipboard(lstEmailImportResults, "Email Import Results");
            buttonPanel.Controls.Add(btnCopyEmailResults);

            var btnCopyAllTabs = new Button();
            btnCopyAllTabs.Text = "Copy All Tabs";
            btnCopyAllTabs.Location = new Point(90, 0);
            btnCopyAllTabs.Size = new Size(100, 25);
            btnCopyAllTabs.Click += (s, e) => CopyAllTabsToClipboard();
            buttonPanel.Controls.Add(btnCopyAllTabs);

            var btnClearEmailResults = new Button();
            btnClearEmailResults.Text = "Clear";
            btnClearEmailResults.Location = new Point(200, 0);
            btnClearEmailResults.Size = new Size(60, 25);
            btnClearEmailResults.Click += (s, e) => lstEmailImportResults.Items.Clear();
            buttonPanel.Controls.Add(btnClearEmailResults);

            lstEmailImportResults = new ListBox();
            lstEmailImportResults.Location = new Point(10, 70);
            lstEmailImportResults.Size = new Size(tabControl.Width - 30, tabControl.Height - 110);
            lstEmailImportResults.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            float fontSize = 10F;
            if (float.TryParse(ConfigurationManager.AppSettings["Form:FontSize"], out float configFontSize))
            {
                fontSize = configFontSize;
            }
            lstEmailImportResults.Font = new Font("Consolas", fontSize);

            lstEmailImportResults.ScrollAlwaysVisible = true;
            lstEmailImportResults.SelectionMode = SelectionMode.MultiExtended;
            tabEmailImport.Controls.Add(lstEmailImportResults);

            ShowDefaultEmailMessage();
        }

        private void ShowDefaultEmailMessage()
        {
            if (lstEmailImportResults != null)
            {
                lstEmailImportResults.Items.Clear();
                lstEmailImportResults.Items.Add("=== ENHANCED EMAIL IMPORT SYSTEM ===");
                lstEmailImportResults.Items.Add("Ready for enhanced email processing");
                lstEmailImportResults.Items.Add("Domain intelligence loaded");
                lstEmailImportResults.Items.Add("Real-time progress tracking enabled");
                lstEmailImportResults.Items.Add("");
                lstEmailImportResults.Items.Add("Click 'Import Emails' to start...");
            }
        }

        private void SetupMkgResultsTab()
        {
            tabMkgResults.Controls.Clear();

            lblMkgStatus = new Label();
            lblMkgStatus.Location = new Point(10, 10);
            lblMkgStatus.Size = new Size(600, 23);
            lblMkgStatus.Text = "MKG injection results will appear here...";
            lblMkgStatus.Font = new Font("Arial", 11F, FontStyle.Bold);
            lblMkgStatus.ForeColor = Color.DarkGreen;
            lblMkgStatus.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            tabMkgResults.Controls.Add(lblMkgStatus);

            var buttonPanel = new Panel();
            buttonPanel.Location = new Point(10, 35);
            buttonPanel.Size = new Size(tabControl.Width - 40, 30);
            buttonPanel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            tabMkgResults.Controls.Add(buttonPanel);

            var btnCopyMkgResults = new Button();
            btnCopyMkgResults.Text = "Copy Tab";
            btnCopyMkgResults.Location = new Point(0, 0);
            btnCopyMkgResults.Size = new Size(80, 25);
            btnCopyMkgResults.Click += (s, e) => CopyListBoxToClipboard(lstMkgResults, "MKG Results");
            buttonPanel.Controls.Add(btnCopyMkgResults);

            var btnCopyAllTabs = new Button();
            btnCopyAllTabs.Text = "Copy All Tabs";
            btnCopyAllTabs.Location = new Point(90, 0);
            btnCopyAllTabs.Size = new Size(100, 25);
            btnCopyAllTabs.Click += (s, e) => CopyAllTabsToClipboard();
            buttonPanel.Controls.Add(btnCopyAllTabs);

            var btnClearMkgResults = new Button();
            btnClearMkgResults.Text = "Clear";
            btnClearMkgResults.Location = new Point(200, 0);
            btnClearMkgResults.Size = new Size(60, 25);
            btnClearMkgResults.Click += (s, e) => lstMkgResults.Items.Clear();
            buttonPanel.Controls.Add(btnClearMkgResults);

            lstMkgResults = new ListBox();
            lstMkgResults.Location = new Point(10, 70);
            lstMkgResults.Size = new Size(tabControl.Width - 30, tabControl.Height - 110);
            lstMkgResults.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            float fontSize = 10F;
            if (float.TryParse(ConfigurationManager.AppSettings["Form:FontSize"], out float configFontSize))
            {
                fontSize = configFontSize;
            }
            lstMkgResults.Font = new Font("Consolas", fontSize);

            lstMkgResults.ScrollAlwaysVisible = true;
            lstMkgResults.SelectionMode = SelectionMode.MultiExtended;
            tabMkgResults.Controls.Add(lstMkgResults);

            ShowDefaultMkgMessage();
        }

        private void ShowDefaultMkgMessage()
        {
            if (lstMkgResults != null)
            {
                lstMkgResults.Items.Clear();
                lstMkgResults.Items.Add("=== MKG TEST RESULTS ===");
                lstMkgResults.Items.Add("");
                lstMkgResults.Items.Add("No MKG tests have been run yet.");
                lstMkgResults.Items.Add("");
                lstMkgResults.Items.Add("To run MKG tests:");
                lstMkgResults.Items.Add("• Click 'MKG' in the menu bar");
                lstMkgResults.Items.Add("• Select 'Mkg Test API'");
                lstMkgResults.Items.Add("");
                lstMkgResults.Items.Add("The test will check:");
                lstMkgResults.Items.Add("✓ Configuration validity");
                lstMkgResults.Items.Add("✓ API connection");
                lstMkgResults.Items.Add("✓ Order system");
                lstMkgResults.Items.Add("✓ Quote system");
                lstMkgResults.Items.Add("✓ Revision system");
                lstMkgResults.Items.Add("");
                lstMkgResults.Items.Add("Results will appear here once tests are run.");
            }
        }

        private void SetupFailedInjectionsTab()
        {
            tabFailedInjections.Controls.Clear();

            lblFailedStatus = new Label();
            lblFailedStatus.Location = new Point(10, 10);
            lblFailedStatus.Size = new Size(600, 23);
            lblFailedStatus.Text = "Failed injection details will appear here...";
            lblFailedStatus.Font = new Font("Arial", 11F, FontStyle.Bold);
            lblFailedStatus.ForeColor = Color.DarkRed;
            lblFailedStatus.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            tabFailedInjections.Controls.Add(lblFailedStatus);

            var buttonPanel = new Panel();
            buttonPanel.Location = new Point(10, 35);
            buttonPanel.Size = new Size(tabControl.Width - 40, 30);
            buttonPanel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            tabFailedInjections.Controls.Add(buttonPanel);

            var btnCopyFailedResults = new Button();
            btnCopyFailedResults.Text = "Copy Tab";
            btnCopyFailedResults.Location = new Point(0, 0);
            btnCopyFailedResults.Size = new Size(80, 25);
            btnCopyFailedResults.Click += (s, e) => CopyListBoxToClipboard(lstFailedInjections, "Failed Injections");
            buttonPanel.Controls.Add(btnCopyFailedResults);

            var btnCopyAllTabs = new Button();
            btnCopyAllTabs.Text = "Copy All Tabs";
            btnCopyAllTabs.Location = new Point(90, 0);
            btnCopyAllTabs.Size = new Size(100, 25);
            btnCopyAllTabs.Click += (s, e) => CopyAllTabsToClipboard();
            buttonPanel.Controls.Add(btnCopyAllTabs);

            var btnClearFailedResults = new Button();
            btnClearFailedResults.Text = "Clear";
            btnClearFailedResults.Location = new Point(200, 0);
            btnClearFailedResults.Size = new Size(60, 25);
            btnClearFailedResults.Click += (s, e) => lstFailedInjections.Items.Clear();
            buttonPanel.Controls.Add(btnClearFailedResults);

            lstFailedInjections = new ListBox();
            lstFailedInjections.Location = new Point(10, 70);
            lstFailedInjections.Size = new Size(tabControl.Width - 30, tabControl.Height - 110);
            lstFailedInjections.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            float fontSize = 10F;
            if (float.TryParse(ConfigurationManager.AppSettings["Form:FontSize"], out float configFontSize))
            {
                fontSize = configFontSize;
            }
            lstFailedInjections.Font = new Font("Consolas", fontSize);

            lstFailedInjections.ScrollAlwaysVisible = true;
            lstFailedInjections.SelectionMode = SelectionMode.MultiExtended;
            tabFailedInjections.Controls.Add(lstFailedInjections);

            ShowDefaultFailedMessage();
        }

        private void ShowDefaultFailedMessage()
        {
            if (lstFailedInjections != null)
            {
                lblFailedStatus.Text = "Ready for test analysis...";
                lblFailedStatus.ForeColor = Color.DarkBlue;

                lstFailedInjections.Items.Clear();
                lstFailedInjections.Items.Add("=== FAILED INJECTIONS ANALYSIS ===");
                lstFailedInjections.Items.Add("");
                lstFailedInjections.Items.Add("No test results available for analysis.");
                lstFailedInjections.Items.Add("");
                lstFailedInjections.Items.Add("This tab will show:");
                lstFailedInjections.Items.Add("• System failures and issues");
                lstFailedInjections.Items.Add("• Configuration problems");
                lstFailedInjections.Items.Add("• API connection failures");
                lstFailedInjections.Items.Add("• Order/Quote/Revision system issues");
                lstFailedInjections.Items.Add("• Success message when all systems work");
                lstFailedInjections.Items.Add("");
                lstFailedInjections.Items.Add("Run an MKG test to see analysis results here.");
            }
        }
        #region Enhanced Duplication Logging Methods
        private void SetupEnhancedDuplicationEventHandlers()
        {
            try
            {
                // Just log MKG duplicates for now - don't change tab status
                DuplicationService.OnManagedMkgDuplicates += (count) =>
                {
                    LogResult($"🔄 Enhanced Detection: {count} MKG duplicates handled automatically");
                };

                Console.WriteLine("✅ Enhanced duplication logging ready");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error setting up enhanced logging: {ex.Message}");
            }
            DuplicationService.OnPriceDiscrepancyDetected += (alert) =>
            {
                LogResult($"🚨 PRICE ALERT: {alert.ArtiCode} - {alert.DiscrepancyPercentage:F1}% difference ({alert.FinancialRisk} risk)");

                // Force RED tab for high financial risks
                if (alert.FinancialRisk == "HIGH")
                {
                    _hasInjectionFailures = true;
                    UpdateTabStatus();
                    LogResult($"🔴 HIGH FINANCIAL RISK DETECTED - Manual review required!");
                }
            };

            // 🚨 CRITICAL BUSINESS ERRORS - Financial Protection
            DuplicationService.OnCriticalBusinessErrors += (count) =>
            {
                LogResult($"🚨 CRITICAL BUSINESS ERRORS: {count} detected - immediate action required");
                _hasInjectionFailures = true; // Force RED tab
                UpdateTabStatus();
            };
        }
        /// <summary>
/// Phase 2: Process email with enhanced price validation
/// Add this call to your email processing workflow
/// </summary>
private EmailDuplicationResult ProcessEmailWithPriceValidation(EmailDetail emailDetail)
{
    try
    {
        Console.WriteLine($"💰 Processing email with price validation: {emailDetail.Subject}");
        
        // Use the enhanced duplication service with price validation
        var result = DuplicationService.ProcessEmailWithBusinessProtection(emailDetail);
        
        // Log price validation results to Email Import tab
        LogPriceValidationResults(result);
        
        return result;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"❌ Error in price validation: {ex.Message}");
        LogResult($"❌ Price validation error: {ex.Message}");
        throw;
    }
}

        /// <summary>
        /// Log price validation results to Email Import tab
        /// </summary>
        private void LogPriceValidationResults(EmailDuplicationResult result)
        {
            try
            {
                if (result?.PriceDiscrepancies?.Any() == true)
                {
                    LogResult($"💰 Price Discrepancies Detected: {result.PriceDiscrepancies.Count}");
                    foreach (var price in result.PriceDiscrepancies)
                    {
                        LogResult($"   ⚠️ {price.ArtiCode}: {price.DiscrepancyPercentage:F1}% difference ({price.FinancialRisk} risk)");
                        LogResult($"   💵 Range: {price.PriceRange}");
                    }
                    LogResult("");
                }

                if (result?.BusinessErrors?.Any() == true)
                {
                    LogResult($"🚨 Critical Business Errors: {result.BusinessErrors.Count}");
                    foreach (var error in result.BusinessErrors)
                    {
                        LogResult($"   ❌ {error.ErrorType}: {error.Description}");
                    }
                    LogResult("");
                }

                // Log cleaned results
                var originalCount = result.OriginalEmail?.Orders?.Count ?? 0;
                var cleanedCount = result.CleanedEmail?.Orders?.Count ?? 0;
                if (originalCount != cleanedCount)
                {
                    LogResult($"🛡️ Financial Protection: {originalCount} → {cleanedCount} items (filtered {originalCount - cleanedCount} risky items)");
                    LogResult("");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error logging price validation: {ex.Message}");
            }
        }
        #endregion

        #region 🎯 Run History System - CLEAN FINAL VERSION

        /// RENAMED: 
        private List<object> CollectEmailResultsForHistory()
        {
            var emailResults = new List<object>();

            try
            {
                if (lstEmailImportResults?.Items.Count > 0)
                {
                    foreach (var item in lstEmailImportResults.Items)
                    {
                        emailResults.Add(item.ToString());
                    }
                }

                Console.WriteLine($"📧 Collected {emailResults.Count} email result items for history");
                return emailResults;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error collecting email results: {ex.Message}");
                return new List<object>();
            }
        }
        /*
        private void SetupEnhancedDuplicationEventHandlers()
        {
            try
            {
                Console.WriteLine("🛡️ Setting up enhanced duplication event handlers...");

                // 🚨 CRITICAL BUSINESS ERRORS - RED tab
                DuplicationService.OnCriticalBusinessErrors += (count) =>
                {
                    try
                    {
                        Console.WriteLine($"🚨 CRITICAL BUSINESS ERRORS DETECTED: {count}");
                        _hasInjectionFailures = true; // Force RED tab
                        UpdateTabStatus();

                        // Update current run data
                        if (_currentRunData != null)
                        {
                            _currentRunData.HasInjectionFailures = true;
                            _currentRunData.TotalFailuresAtCompletion += count;
                        }

                        // Log critical errors in Failed Injections tab
                        LogCriticalBusinessErrors();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error handling critical business errors: {ex.Message}");
                    }
                };

                // 💰 PRICE DISCREPANCY ALERTS - RED tab  
                DuplicationService.OnPriceDiscrepancyDetected += (alert) =>
                {
                    try
                    {
                        Console.WriteLine($"💰 PRICE DISCREPANCY: {alert.ArtiCode} - {alert.DiscrepancyPercentage:F1}% difference");
                        _hasInjectionFailures = true; // Force RED tab for financial risks
                        UpdateTabStatus();

                        // Update current run data
                        if (_currentRunData != null)
                        {
                            _currentRunData.HasInjectionFailures = true;
                            _currentRunData.TotalFailuresAtCompletion++;
                        }

                        // Log price discrepancy in Failed Injections tab
                        LogPriceDiscrepancy(alert);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error handling price discrepancy: {ex.Message}");
                    }
                };

                // 🔄 MANAGED MKG DUPLICATES - YELLOW tab
                DuplicationService.OnManagedMkgDuplicates += (count) =>
                {
                    try
                    {
                        Console.WriteLine($"🔄 MANAGED DUPLICATES DETECTED: {count}");
                        _hasMkgDuplicatesDetected = true; // Force YELLOW tab (if no critical errors)
                        _totalDuplicatesInCurrentRun += count;
                        UpdateTabStatus();

                        // Update current run data
                        if (_currentRunData != null)
                        {
                            _currentRunData.HasMkgDuplicatesDetected = true;
                            _currentRunData.TotalDuplicatesDetected += count;
                        }

                        // Log managed duplicates
                        LogManagedDuplicates(count);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error handling managed duplicates: {ex.Message}");
                    }
                };

                // 📧 EMAIL-LEVEL DUPLICATES - Information only
                DuplicationService.OnEmailLevelDuplicatesDetected += (count) =>
                {
                    try
                    {
                        Console.WriteLine($"📧 EMAIL-LEVEL DUPLICATES: {count} (HTML vs PDF conflicts)");
                        LogResult($"🔍 Email-level duplicates detected: {count} (HTML vs PDF extraction conflicts)");

                        // These don't change tab status but are logged for transparency
                        if (_enhancedProgress != null)
                        {
                            _enhancedProgress.IncrementEmailDuplicates();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error handling email-level duplicates: {ex.Message}");
                    }
                };

                // 🔄 CROSS-EMAIL DUPLICATES - Information only  
                DuplicationService.OnCrossEmailDuplicatesDetected += (count) =>
                {
                    try
                    {
                        Console.WriteLine($"🔄 CROSS-EMAIL DUPLICATES: {count} (filtered automatically)");
                        LogResult($"🔄 Cross-email duplicates filtered: {count} (already processed in previous emails)");

                        // These don't change tab status but are logged for transparency
                        if (_enhancedProgress != null)
                        {
                            _enhancedProgress.IncrementEmailDuplicates();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error handling cross-email duplicates: {ex.Message}");
                    }
                };

                Console.WriteLine("✅ Enhanced duplication event handlers setup completed");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error setting up duplication event handlers: {ex.Message}");
            }
        }
        */
        /// <summary>
        /// RENAMED: Get current MKG results for saving to run history
        /// </summary>
        private List<object> CollectMkgResultsForHistory()
        {
            var mkgResults = new List<object>();

            try
            {
                if (lstMkgResults?.Items.Count > 0)
                {
                    foreach (var item in lstMkgResults.Items)
                    {
                        mkgResults.Add(item.ToString());
                    }
                }

                Console.WriteLine($"📊 Collected {mkgResults.Count} MKG result items for history");
                return mkgResults;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error collecting MKG results: {ex.Message}");
                return new List<object>();
            }
        }

        /// <summary>
        /// RENAMED: Get current error results for saving to run history
        /// </summary>
        private List<object> CollectErrorResultsForHistory()
        {
            var errorResults = new List<object>();

            try
            {
                if (lstFailedInjections?.Items.Count > 0)
                {
                    foreach (var item in lstFailedInjections.Items)
                    {
                        var itemText = item.ToString();
                        // Only save actual error content, not the default messages
                        if (!itemText.Contains("Ready for") &&
                            !itemText.Contains("No test results") &&
                            !itemText.Contains("Error tracking active"))
                        {
                            errorResults.Add(itemText);
                        }
                    }
                }

                Console.WriteLine($"❌ Collected {errorResults.Count} error result items for history");
                return errorResults;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error collecting error results: {ex.Message}");
                return new List<object>();
            }
        }
        private void InitializeRunHistory()
        {
            try
            {
                _runHistoryManager = new RunHistoryManager();
                _currentRunData = new AutomationRunData();

                // Create dropdown
                cmbRunHistory = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    DisplayMember = "DisplayText",
                    ValueMember = "RunId",
                    Font = new Font("Segoe UI", 9F),
                    Width = 450,
                    Location = new Point(12, 27),
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                };

                // Add proper event handler
                cmbRunHistory.SelectedIndexChanged += CmbRunHistory_SelectedIndexChanged;

                this.Controls.Add(cmbRunHistory);

                // Move TabControl down
                tabControl.Location = new Point(12, 60);
                tabControl.Size = new Size(tabControl.Size.Width, tabControl.Size.Height - 25);

                RefreshRunHistory();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Run history init error: {ex.Message}");
            }
        }

        // 🎯 Event handler that loads historical data
        private async void CmbRunHistory_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cmbRunHistory.SelectedItem is RunHistoryItem selectedRun)
                {
                    Console.WriteLine($"🎯 Selected run: {selectedRun.RunId} (IsCurrentRun: {selectedRun.IsCurrentRun})");

                    if (selectedRun.IsCurrentRun)
                    {
                        // Current run - restore current data
                        _isLoadingHistoricalRun = false;
                        RestoreCurrentRunView();
                        Console.WriteLine("✅ Switched to current run view");
                    }
                    else
                    {
                        // Historical run - load the data
                        _isLoadingHistoricalRun = true;
                        Console.WriteLine($"📚 Loading historical run: {selectedRun.StartTime:yyyy-MM-dd HH:mm}");
                        await LoadHistoricalRunData(selectedRun.RunId);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in run selection: {ex.Message}");
                MessageBox.Show($"Error loading selected run: {ex.Message}", "Run History Error",
                               MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // 🎯 Load historical run data
        private async Task LoadHistoricalRunData(Guid runId)
        {
            try
            {
                Console.WriteLine($"📚 Loading historical run data for: {runId}");

                // Load the complete run data
                var historicalRunData = _runHistoryManager.LoadRunData(runId);

                if (historicalRunData == null)
                {
                    Console.WriteLine("⚠️ No detailed data found for this run");

                    // Show basic info from RunHistoryItem instead
                    var runInfo = _runHistoryManager.GetAllRuns().FirstOrDefault(r => r.RunId == runId);
                    if (runInfo != null)
                    {
                        DisplayBasicHistoricalInfo(runInfo);
                    }
                    else
                    {
                        MessageBox.Show("No data available for this run.", "Run History",
                                       MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    return;
                }

                Console.WriteLine($"✅ Loaded historical data");
                Console.WriteLine($"   EmailResults: {historicalRunData.EmailResults?.Count ?? 0}");
                Console.WriteLine($"   MkgResults: {historicalRunData.MkgResults?.Count ?? 0}");
                Console.WriteLine($"   ErrorResults: {historicalRunData.ErrorResults?.Count ?? 0}");

                // Update the UI with historical data using the EXISTING methods
                DisplayHistoricalRunData(historicalRunData);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error loading historical run data: {ex.Message}");
                MessageBox.Show($"Error loading run data: {ex.Message}", "Run History Error",
                               MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // Keep the UpdateFailedInjectionsTabHeaderColor method but only call it:
        // 1. After MKG injection completion
        // 2. When displaying truly completed runs
        private void UpdateFailedInjectionsTabHeaderColor(object dataSource)
        {
            try
            {
                bool hasErrors = false;
                bool hasDuplicates = false;

                // Determine error status from data source
                if (dataSource is AutomationRunData runData)
                {
                    hasErrors = runData.HasInjectionFailures ||
                               (runData.ErrorResults?.Count ?? 0) > 0 ||
                               runData.TotalFailuresAtCompletion > 0;
                    hasDuplicates = runData.HasMkgDuplicatesDetected ||
                                  runData.TotalDuplicatesDetected > 0;
                }
                else if (dataSource is RunHistoryItem runInfo)
                {
                    hasErrors = runInfo.HasInjectionFailures || runInfo.TotalFailuresAtCompletion > 0;
                    // Check for duplicates in settings
                    if (runInfo.Settings?.ContainsKey("TotalDuplicates") == true)
                    {
                        if (int.TryParse(runInfo.Settings["TotalDuplicates"].ToString(), out int dups))
                        {
                            hasDuplicates = dups > 0;
                        }
                    }
                }

                // Find the Failed Injections tab and update ONLY its color
                foreach (TabPage tab in tabControl.TabPages)
                {
                    if (tab.Text.Contains("Failed") || tab.Text.Contains("Error"))
                    {
                        if (hasErrors)
                        {
                            tab.BackColor = Color.Red;
                            tab.ForeColor = Color.White;
                        }
                        else if (hasDuplicates)
                        {
                            tab.BackColor = Color.Orange;
                            tab.ForeColor = Color.White;
                        }
                        else
                        {
                            tab.BackColor = Color.Green;
                            tab.ForeColor = Color.White;
                        }

                        Console.WriteLine($"🎨 Updated Failed Injections tab color: Errors={hasErrors}, Duplicates={hasDuplicates}");
                        break; // Only update the Failed Injections tab
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating tab color: {ex.Message}");
            }
        }
        // 🎯 Update status labels for basic history
        private void UpdateStatusLabelsForBasicHistory(RunHistoryItem runInfo)
        {
            try
            {
                if (lblMkgStatus != null)
                {
                    var totalFound = runInfo.OrdersFound + runInfo.QuotesFound + runInfo.RevisionsFound;
                    lblMkgStatus.Text = $"📊 Historical: {totalFound} items found, {runInfo.ErrorsEncountered} errors";
                    lblMkgStatus.ForeColor = runInfo.ErrorsEncountered > 0 ? Color.Orange : Color.Green;
                }

                if (lblFailedStatus != null)
                {
                    if (runInfo.HasInjectionFailures || runInfo.TotalFailuresAtCompletion > 0)
                    {
                        lblFailedStatus.Text = $"❌ Historical: {runInfo.TotalFailuresAtCompletion} failed injections";
                        lblFailedStatus.ForeColor = Color.Red;
                    }
                    else
                    {
                        lblFailedStatus.Text = "✅ Historical: No failed injections";
                        lblFailedStatus.ForeColor = Color.Green;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating basic history status: {ex.Message}");
            }
        }

        // 🎯 Update status labels for detailed history
        private void UpdateStatusLabelsForDetailedHistory(AutomationRunData historicalData)
        {
            try
            {
                if (lblMkgStatus != null)
                {
                    var mkgCount = historicalData.MkgResults?.Count ?? 0;
                    var errorCount = historicalData.ErrorResults?.Count ?? 0;
                    lblMkgStatus.Text = $"📊 Historical: {mkgCount} MKG results, {errorCount} errors";
                    lblMkgStatus.ForeColor = errorCount > 0 ? Color.Orange : Color.Green;
                }

                if (lblFailedStatus != null)
                {
                    var totalFailures = historicalData.TotalFailuresAtCompletion;
                    if (historicalData.HasInjectionFailures || totalFailures > 0)
                    {
                        lblFailedStatus.Text = $"❌ Historical: {totalFailures} failed injections";
                        lblFailedStatus.ForeColor = Color.Red;
                    }
                    else
                    {
                        lblFailedStatus.Text = "✅ Historical: No failed injections";
                        lblFailedStatus.ForeColor = Color.Green;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating detailed history status: {ex.Message}");
            }
        }

        // 🎯 Restore current run view
        private void RestoreCurrentRunView()
        {
            try
            {
                Console.WriteLine("🔄 Restoring current run view...");

                // Reset any historical indicators in the status labels
                if (lblMkgStatus != null && lblMkgStatus.Text.Contains("Historical"))
                {
                    lblMkgStatus.Text = "Ready for current run";
                    lblMkgStatus.ForeColor = Color.Black;
                }

                if (lblFailedStatus != null && lblFailedStatus.Text.Contains("Historical"))
                {
                    lblFailedStatus.Text = "Ready for current run";
                    lblFailedStatus.ForeColor = Color.Black;
                }

                // 🔧 FIXED: Reset tab coloring to current run state
                if (tabControl != null)
                {
                    // Reset to current run status (this will use current _hasInjectionFailures and _hasMkgDuplicatesDetected)
                    tabControl.Invalidate();
                    tabControl.Update();
                }

                Console.WriteLine("✅ Current run view restored");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error restoring current run view: {ex.Message}");
            }
        }

        private void RefreshRunHistory()
        {
            try
            {
                cmbRunHistory.Items.Clear();
                var runs = _runHistoryManager.GetAllRuns();
                foreach (var run in runs)
                    cmbRunHistory.Items.Add(run);

                var currentRun = runs.FirstOrDefault(r => r.IsCurrentRun);
                if (currentRun != null)
                    cmbRunHistory.SelectedItem = currentRun;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Refresh error: {ex.Message}");
            }
        }

        // 🎯 FIXED: Only update Failed Injections tab color (not whole form)
        private void UpdateTabColorForFailedInjectionsOnly(object dataSource)
        {
            try
            {
                bool hasErrors = false;
                bool hasDuplicates = false;

                // Determine error status from data source
                if (dataSource is AutomationRunData runData)
                {
                    hasErrors = runData.HasInjectionFailures ||
                               (runData.ErrorResults?.Count ?? 0) > 0 ||
                               runData.TotalFailuresAtCompletion > 0;
                    hasDuplicates = runData.HasMkgDuplicatesDetected ||
                                  runData.TotalDuplicatesDetected > 0;
                }
                else if (dataSource is RunHistoryItem runInfo)
                {
                    hasErrors = runInfo.HasInjectionFailures || runInfo.TotalFailuresAtCompletion > 0;
                    // Check for duplicates in settings
                    if (runInfo.Settings?.ContainsKey("TotalDuplicates") == true)
                    {
                        if (int.TryParse(runInfo.Settings["TotalDuplicates"].ToString(), out int dups))
                        {
                            hasDuplicates = dups > 0;
                        }
                    }
                }

                // Find the Failed Injections tab and update ONLY its color
                foreach (TabPage tab in tabControl.TabPages)
                {
                    if (tab.Text.Contains("Failed") || tab.Text.Contains("Error"))
                    {
                        if (hasErrors)
                        {
                            tab.BackColor = Color.Red;
                            tab.ForeColor = Color.White;
                        }
                        else if (hasDuplicates)
                        {
                            tab.BackColor = Color.Orange;
                            tab.ForeColor = Color.White;
                        }
                        else
                        {
                            tab.BackColor = Color.Green;
                            tab.ForeColor = Color.White;
                        }

                        Console.WriteLine($"🎨 Updated Failed Injections tab color: Errors={hasErrors}, Duplicates={hasDuplicates}");
                        break; // Only update the Failed Injections tab
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating tab color: {ex.Message}");
            }
        }

        // 🎯 FIXED: Display historical email results
        private void DisplayHistoricalEmailResults(AutomationRunData historicalData)
        {
            try
            {
                if (lstEmailImportResults == null) return;

                lstEmailImportResults.Items.Add("=== HISTORICAL EMAIL IMPORT RESULTS ===");
                lstEmailImportResults.Items.Add($"📅 Run: {historicalData.RunInfo?.StartTime:yyyy-MM-dd HH:mm:ss}");
                lstEmailImportResults.Items.Add("");

                if (historicalData.EmailResults?.Any() == true)
                {
                    lstEmailImportResults.Items.Add($"📧 Email Results Found: {historicalData.EmailResults.Count}");
                    lstEmailImportResults.Items.Add("");

                    // Display each email result
                    int count = 0;
                    foreach (var result in historicalData.EmailResults) // Show first 10
                    {
                        count++;
                        lstEmailImportResults.Items.Add(result?.ToString() ?? "No data");
                    }
                }
                else
                {
                    lstEmailImportResults.Items.Add("📧 No email results stored for this run");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error displaying historical email results: {ex.Message}");
            }
        }

        // 🎯 FIXED: Display historical MKG results
        private void DisplayHistoricalMkgResults(AutomationRunData historicalData)
        {
            try
            {
                if (lstMkgResults == null) return;

                lstMkgResults.Items.Add("=== HISTORICAL MKG INJECTION RESULTS ===");
                lstMkgResults.Items.Add($"📅 Run: {historicalData.RunInfo?.StartTime:yyyy-MM-dd HH:mm:ss}");
                lstMkgResults.Items.Add("");

                if (historicalData.MkgResults?.Any() == true)
                {
                    lstMkgResults.Items.Add($"📊 MKG Results Found: {historicalData.MkgResults.Count}");
                    lstMkgResults.Items.Add("");

                    foreach (var result in historicalData.MkgResults)
                    {
                        lstMkgResults.Items.Add($"• {result?.ToString() ?? "No data"}");
                    }
                }
                else
                {
                    lstMkgResults.Items.Add("📊 No MKG injection results stored for this run");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error displaying historical MKG results: {ex.Message}");
            }
        }

        // 🎯 FIXED: Display historical failed injections
        private void DisplayHistoricalFailedInjections(AutomationRunData historicalData)
        {
            try
            {
                if (lstFailedInjections == null) return;

                lstFailedInjections.Items.Add("=== HISTORICAL FAILED INJECTIONS ===");
                lstFailedInjections.Items.Add($"📅 Run: {historicalData.RunInfo?.StartTime:yyyy-MM-dd HH:mm:ss}");
                lstFailedInjections.Items.Add("");

                // Check for failures
                var hasFailures = historicalData.HasInjectionFailures ||
                                 historicalData.TotalFailuresAtCompletion > 0 ||
                                 (historicalData.ErrorResults?.Count ?? 0) > 0;

                if (!hasFailures)
                {
                    lstFailedInjections.Items.Add("✅ NO INJECTION FAILURES in this historical run!");
                    lstFailedInjections.Items.Add("");
                    lstFailedInjections.Items.Add("All injections completed successfully.");
                    return;
                }

                // Show failure details
                lstFailedInjections.Items.Add($"❌ Injection failures detected: {historicalData.TotalFailuresAtCompletion}");
                lstFailedInjections.Items.Add($"📊 Error results stored: {historicalData.ErrorResults?.Count ?? 0}");
                lstFailedInjections.Items.Add("");

                if (historicalData.ErrorResults?.Any() == true)
                {
                    lstFailedInjections.Items.Add("Error Details:");
                    foreach (var error in historicalData.ErrorResults)
                    {
                        lstFailedInjections.Items.Add($"• {error?.ToString() ?? "Unknown error"}");
                    }
                }

                // Show duplicate information if available
                if (historicalData.HasMkgDuplicatesDetected)
                {
                    lstFailedInjections.Items.Add("");
                    lstFailedInjections.Items.Add($"🔄 Duplicates detected: {historicalData.TotalDuplicatesDetected}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error displaying historical failed injections: {ex.Message}");
            }
        }

        // 🎯 NEW: Update status labels for basic history


        // 🎯 NEW: Update status labels for detailed history

        #endregion

        private async Task UpdateUIWithHistoricalData(AutomationRunData historicalData)
        {
            try
            {
                Console.WriteLine("🔄 Updating UI with historical data...");

                // Clear current displays
                lstEmailImportResults?.Items.Clear();
                lstMkgResults?.Items.Clear();
                lstFailedInjections?.Items.Clear();

                // Update Email Import Results Tab
                if (historicalData.EmailResults != null && historicalData.EmailResults.Any())
                {
                    Console.WriteLine($"📧 Displaying {historicalData.EmailResults.Count} email results");
                    DisplayHistoricalEmailResults(historicalData.EmailResults);
                }

                // Update MKG Results Tab
                if (historicalData.MkgResults != null && historicalData.MkgResults.Any())
                {
                    Console.WriteLine("📊 Displaying MKG injection results");
                    DisplayHistoricalMkgResults(historicalData.MkgResults);
                }

                // Update Failed Injections Tab
                if (historicalData.ErrorResults != null && historicalData.ErrorResults.Any())
                {
                    Console.WriteLine("❌ Displaying failed injections");
                    DisplayHistoricalErrorResults(historicalData.ErrorResults);
                }

                // Update status labels
                UpdateStatusLabelsForHistoricalData(historicalData);

                Console.WriteLine("✅ UI updated with historical data");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating UI with historical data: {ex.Message}");
            }
        }

        // 🎯 NEW: Display historical email results
        private void DisplayHistoricalEmailResults(List<object> emailResults)
        {
            try
            {
                if (lstEmailImportResults == null) return;

                lstEmailImportResults.Items.Clear();
                lstEmailImportResults.Items.Add("=== HISTORICAL EMAIL IMPORT RESULTS ===");
                lstEmailImportResults.Items.Add($"📅 Viewing historical data");
                lstEmailImportResults.Items.Add("");

                // Try to cast email results to proper format
                foreach (var result in emailResults)
                {
                    try
                    {
                        // Convert the object to a displayable format
                        var resultText = result?.ToString() ?? "Unknown result";
                        lstEmailImportResults.Items.Add(resultText);
                    }
                    catch (Exception ex)
                    {
                        lstEmailImportResults.Items.Add($"❌ Error displaying result: {ex.Message}");
                    }
                }

                lstEmailImportResults.Items.Add("");
                lstEmailImportResults.Items.Add($"📊 Total historical email results: {emailResults.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error displaying historical email results: {ex.Message}");
            }
        }

        // 🎯 NEW: Display historical MKG results
        private void DisplayHistoricalMkgResults(List<object> mkgResults)
        {
            try
            {
                if (lstMkgResults == null) return;

                lstMkgResults.Items.Clear();
                lstMkgResults.Items.Add("=== HISTORICAL MKG INJECTION RESULTS ===");
                lstMkgResults.Items.Add($"📅 Viewing historical MKG data");
                lstMkgResults.Items.Add("");

                foreach (var result in mkgResults)
                {
                    try
                    {
                        var resultText = result?.ToString() ?? "Unknown MKG result";
                        lstMkgResults.Items.Add(resultText);
                    }
                    catch (Exception ex)
                    {
                        lstMkgResults.Items.Add($"❌ Error displaying MKG result: {ex.Message}");
                    }
                }

                lstMkgResults.Items.Add("");
                lstMkgResults.Items.Add($"📊 Total historical MKG results: {mkgResults.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error displaying historical MKG results: {ex.Message}");
            }
        }

        // 🎯 NEW: Display historical error results
        private void DisplayHistoricalErrorResults(List<object> errorResults)
        {
            try
            {
                if (lstFailedInjections == null) return;

                lstFailedInjections.Items.Clear();
                lstFailedInjections.Items.Add("=== HISTORICAL ERROR RESULTS ===");
                lstFailedInjections.Items.Add($"📅 Viewing historical error data");
                lstFailedInjections.Items.Add("");

                if (!errorResults.Any())
                {
                    lstFailedInjections.Items.Add("✅ No errors in this historical run!");
                    return;
                }

                foreach (var error in errorResults)
                {
                    try
                    {
                        var errorText = error?.ToString() ?? "Unknown error";
                        lstFailedInjections.Items.Add(errorText);
                    }
                    catch (Exception ex)
                    {
                        lstFailedInjections.Items.Add($"❌ Error displaying error result: {ex.Message}");
                    }
                }

                lstFailedInjections.Items.Add("");
                lstFailedInjections.Items.Add($"📊 Total historical errors: {errorResults.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error displaying historical error results: {ex.Message}");
            }
        }

        // 🎯 NEW: Update status labels for historical data - FIXED TO USE CORRECT PROPERTIES
        private void UpdateStatusLabelsForHistoricalData(AutomationRunData historicalData)
        {
            try
            {
                // Update MKG status based on available data
                if (lblMkgStatus != null)
                {
                    var mkgResultsCount = historicalData.MkgResults?.Count ?? 0;
                    var errorResultsCount = historicalData.ErrorResults?.Count ?? 0;
                    var successfulResults = Math.Max(0, mkgResultsCount - errorResultsCount);

                    lblMkgStatus.Text = $"📊 Historical: {successfulResults} successful, {errorResultsCount} failed";
                    lblMkgStatus.ForeColor = errorResultsCount > 0 ? Color.Orange : Color.Green;
                }

                // Update Failed status
                if (lblFailedStatus != null)
                {
                    var totalErrors = historicalData.ErrorResults?.Count ?? 0;
                    var hasInjectionFailures = historicalData.HasInjectionFailures;
                    var totalFailures = historicalData.TotalFailuresAtCompletion;

                    if (totalErrors > 0 || hasInjectionFailures || totalFailures > 0)
                    {
                        lblFailedStatus.Text = $"❌ Historical: {Math.Max(totalErrors, totalFailures)} failed injections";
                        lblFailedStatus.ForeColor = Color.Red;
                    }
                    else
                    {
                        lblFailedStatus.Text = "✅ Historical: No failed injections";
                        lblFailedStatus.ForeColor = Color.Green;
                    }
                }

                // Update tab colors based on historical data
                UpdateTabColorsForHistoricalData(historicalData);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating status labels: {ex.Message}");
            }
        }

        // 🎯 NEW: Update tab colors for historical data
        private void UpdateTabColorsForHistoricalData(AutomationRunData historicalData)
        {
            try
            {
                // Determine tab color based on historical data
                var hasErrors = historicalData.HasInjectionFailures ||
                               (historicalData.ErrorResults?.Count ?? 0) > 0 ||
                               historicalData.TotalFailuresAtCompletion > 0;

                var hasDuplicates = historicalData.HasMkgDuplicatesDetected ||
                                  historicalData.TotalDuplicatesDetected > 0;

                // Apply the same tab color logic as current runs
                if (hasErrors)
                {
                    // Red for errors
                    ApplyTabColors(Color.Red, Color.White);
                }
                else if (hasDuplicates)
                {
                    // Orange for duplicates
                    ApplyTabColors(Color.Orange, Color.White);
                }
                else
                {
                    // Green for success
                    ApplyTabColors(Color.Green, Color.White);
                }

                Console.WriteLine($"🎨 Applied historical tab colors: Errors={hasErrors}, Duplicates={hasDuplicates}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating tab colors: {ex.Message}");
            }
        }

        // 🎯 Helper method to apply tab colors (if not already present)
        private void ApplyTabColors(Color backColor, Color foreColor)
        {
            try
            {
                // Apply to all tabs in the tab control
                foreach (TabPage tab in tabControl.TabPages)
                {
                    tab.BackColor = backColor;
                    tab.ForeColor = foreColor;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error applying tab colors: {ex.Message}");
            }
        }


        // Replace your TabControl_DrawItem method with this version:
        private void TabControl_DrawItem(object sender, DrawItemEventArgs e)
        {
            try
            {
                TabControl tc = (TabControl)sender;
                TabPage tp = tc.TabPages[e.Index];
                Rectangle tabRect = tc.GetTabRect(e.Index);

                // Default colors
                Color bgColor = SystemColors.Control;
                Color textColor = SystemColors.ControlText;

                // Check if this is the Failed Injections tab
                if (tp.Text.Contains("Failed"))
                {
                    var status = GetCurrentTabStatus();
                    switch (status)
                    {
                        case TabStatus.Red:
                            bgColor = Color.LightCoral;
                            textColor = Color.DarkRed;
                            break;
                        case TabStatus.Yellow:
                            bgColor = Color.LightYellow;
                            textColor = Color.DarkOrange;
                            break;
                        case TabStatus.Green:
                            bgColor = Color.LightGreen;
                            textColor = Color.DarkGreen;
                            break;
                    }
                }

                // Draw background
                using (var brush = new SolidBrush(bgColor))
                    e.Graphics.FillRectangle(brush, tabRect);

                // Draw text
                TextRenderer.DrawText(e.Graphics, tp.Text, tc.Font, tabRect, textColor,
                                     TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Tab draw error: {ex.Message}");
                // Fallback to default drawing
                e.DrawBackground();
            }
        }
        private void OnDuplicatesDetected(int duplicateCount)
        {
            try
            {
                Console.WriteLine($"🔥 OnDuplicatesDetected triggered with count: {duplicateCount}");

                // Update the tracking variables (these exist in Elcotec class)
                _hasMkgDuplicatesDetected = true;
                _totalDuplicatesInCurrentRun += duplicateCount;

                // Update the tab status immediately (this method exists in Elcotec class)
                UpdateTabStatus();

                // Update the current run data if available (this exists in Elcotec class)
                if (_currentRunData != null)
                {
                    _currentRunData.HasMkgDuplicatesDetected = true;
                    _currentRunData.TotalDuplicatesDetected += duplicateCount;
                    _currentRunData.LastDuplicateDetectionUpdate = DateTime.Now;
                }

                Console.WriteLine($"✅ Duplicates processed: Total={_totalDuplicatesInCurrentRun}, Tab should be YELLOW");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in OnDuplicatesDetected: {ex.Message}");
            }
        }

        private async void btnEnhancedEmailImport_Click(object sender, EventArgs e)
        {
            if (_isProcessing)
            {
                MessageBox.Show("Email import already in progress.", "Import In Progress",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                _isProcessing = true;
                _processingStartTime = DateTime.Now;

                // 🔥 CRITICAL: Reset all status tracking for new run
                _hasInjectionFailures = false;
                _hasMkgDuplicatesDetected = false;
                _totalDuplicatesInCurrentRun = 0;
                Console.WriteLine("🟢 NEW RUN STARTED - All status flags reset to clean state");

                // 🔥 CRITICAL: Update tab to green at start
                UpdateTabStatus();

                // 🎯 NEW: Start new run history tracking
                _currentRunId = _runHistoryManager.CreateNewRun();
                _currentRunData = new AutomationRunData
                {
                    RunInfo = _runHistoryManager.GetCurrentRun(),
                    HasInjectionFailures = false,
                    HasMkgDuplicatesDetected = false,
                    TotalDuplicatesDetected = 0,
                    TotalFailuresAtCompletion = 0
                };

                Console.WriteLine($"🆕 Created new run: {_currentRunId}");
                RefreshRunHistory();

                // ✅ CRITICAL: START THE ENHANCED PROGRESS MANAGER
                Console.WriteLine("🚀 Starting ENHANCED email import with floating stats...");
                _enhancedProgress?.StartNewSession("Enhanced Email Import", 100);

                // Clear results and initialize
                ClearAllResults();
                UpdateProgress(0, 100, "Starting enhanced email import...");

                Console.WriteLine("🚀 Starting ENHANCED email import (Orders + Quotes + Revisions)...");

                // Check domain mappings
                try
                {
                    var domainMappings = EnhancedDomainProcessor.GetAllDomainMappings();
                    LogResult($"🔧 Domain processor loaded with {domainMappings.Count} customer mappings");

                    var weirDomains = domainMappings.Where(d => d.Key.Contains("weir")).ToList();
                    if (weirDomains.Any())
                    {
                        LogResult($"⭐ Found {weirDomains.Count} Weir domain configurations");
                        foreach (var weir in weirDomains)
                        {
                            LogResult($"   🏢 {weir.Key} → {weir.Value?.ToString() ?? "Unknown"}");
                        }
                    }
                }
                catch (Exception domainEx)
                {
                    LogResult($"⚠️ Domain processor warning: {domainEx.Message}");
                }

                // ✅ CREATE ENHANCED PROGRESS CALLBACK
                var enhancedProgress = new EnhancedProgressCallback(_enhancedProgress);

                // Import emails (this now includes extraction with HTML parsers)
                _enhancedProgress?.UpdateProgress(10, 100, "Importing emails and extracting business content...");
                var importSummary = await EmailWorkFlowService.ImportEmailsAsync(_enhancedProgress);

                // ✅ FIXED: Update statistics ONLY for emails processed (not for individual items)
                if (importSummary?.EmailDetails?.Count > 0)
                {
                    foreach (var email in importSummary.EmailDetails)
                    {
                        _enhancedProgress?.IncrementEmailsProcessed();

                        // 🎯 NEW: Update run progress
                        if (_currentRunId != Guid.Empty)
                        {
                            var emailsProcessed = _enhancedProgress?.EmailsProcessed ?? 0;
                            var ordersFound = _enhancedProgress?.OrdersFound ?? 0;
                            var quotesFound = _enhancedProgress?.QuotesFound ?? 0;
                            var revisionsFound = _enhancedProgress?.RevisionsFound ?? 0;

                            _runHistoryManager.UpdateRunProgress(_currentRunId, emailsProcessed, ordersFound,
                                quotesFound, revisionsFound, importSummary.EmailDetails.Count);
                            RefreshRunHistory();
                        }

                        await Task.Delay(1); // Small delay to see the counter update
                    }
                }

                // Display email results first
                DisplayEmailImportResults(importSummary);

                // Switch to MKG tab for injection
                if (tabControl != null && tabControl.TabPages.Count > 1)
                {
                    tabControl.SelectedIndex = 1; // MKG Results tab
                }

                // ENHANCED INJECTION: Orders + Quotes + Revisions
                _enhancedProgress?.UpdateProgress(70, 100, "Starting enhanced MKG injection (Orders + Quotes + Revisions)...");

                // Process MKG injection with enhanced progress
                await ProcessEnhancedMkgInjection(importSummary);

                // Final completion
                var totalTime = DateTime.Now - _processingStartTime;
                _enhancedProgress?.CompleteOperation($"Enhanced email import completed in {totalTime.TotalSeconds:F1}s!");

                // 🎯 NEW: Save final run data WITH ACTUAL CONTENT
                if (_currentRunData != null)
                {
                    PersistCurrentRunDataNow(); // 🔧 This replaces the TODO comments
                }

                RefreshRunHistory();

                LogResult($"🎉 Enhanced email import completed in {totalTime.TotalSeconds:F1} seconds");
                Console.WriteLine("✅ ENHANCED workflow finished!");
            }
            catch (Exception ex)
            {
                _isProcessing = false;
                Console.WriteLine($"❌ Enhanced button click error: {ex.Message}");
                LogResult($"❌ Error in enhanced email processing: {ex.Message}");

                // ✅ HANDLE ERRORS WITH ENHANCED PROGRESS
                _enhancedProgress?.FailOperation(ex.Message);

                // 🎯 NEW: Complete run with error status
                if (_currentRunId != Guid.Empty)
                {
                    _runHistoryManager.CompleteRun(_currentRunId, "Failed");
                    RefreshRunHistory();
                }

                UpdateProgress(0, 100, $"Error: {ex.Message}");

                MessageBox.Show($"Error during enhanced email processing: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _isProcessing = false;
                UpdateProgress(-1, -1, "Ready"); // Hide progress bar
            }
        }
        private void InitializeEnhancedProgress()
        {
            try
            {
                _enhancedProgress = new EnhancedProgressManager(
                    this,
                    tabControl,
                    toolStripProgressBar,
                    toolStripStatusLabel,
                    toolStripProgressLabel
                );

                EnhancedDomainProcessor.LoadDomainMappingsFromConfig();
                Console.WriteLine("Enhanced progress manager initialized successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Enhanced progress initialization warning: {ex.Message}");
                _enhancedProgress = null;
            }
        }
        #endregion

        #region Enhanced Button Click Handlers

        /// <summary>
        /// Main Email Import button handler - ENHANCED with Orders + Quotes + Revisions
        /// </summary>
        private async void btnProcessEmails_Click(object sender, EventArgs e)
        {
            await ProcessEnhancedEmailImport();
        }

        /// <summary>
        /// Menu Email Import handler - ENHANCED with Orders + Quotes + Revisions
        /// </summary>
        private async void Email_Import_Tick(object sender, EventArgs e)
        {
            await ProcessEnhancedEmailImport();
        }

        /// <summary>
        /// CLEANED ENHANCED EMAIL IMPORT METHOD - Handles Orders + Quotes + Revisions
        /// </summary>
        /// <summary>
        /// ENHANCED EMAIL IMPORT METHOD with Run History Integration
        /// </summary>
        private async Task ProcessEnhancedEmailImport()
        {
            if (_isProcessing)
            {
                MessageBox.Show("Email import already in progress.", "Import In Progress",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                _isProcessing = true;
                _processingStartTime = DateTime.Now;

                // Reset status tracking
                _hasInjectionFailures = false;
                _hasMkgDuplicatesDetected = false;
                _totalDuplicatesInCurrentRun = 0;
                UpdateTabStatus();

                // Start new run history tracking
                _currentRunId = _runHistoryManager.CreateNewRun();
                _currentRunData = new AutomationRunData
                {
                    RunInfo = _runHistoryManager.GetCurrentRun(),
                    HasInjectionFailures = false,
                    HasMkgDuplicatesDetected = false,
                    TotalDuplicatesDetected = 0,
                    TotalFailuresAtCompletion = 0
                };

                RefreshRunHistory();

                // START THE ENHANCED PROGRESS MANAGER
                _enhancedProgress?.StartOperation("Enhanced Email Import", 100);

                // Clear results and initialize
                ClearAllResults();
                UpdateProgress(0, 100, "Starting enhanced email import...");

                // Domain mappings check with real-time spacing
                try
                {
                    var domainMappings = EnhancedDomainProcessor.GetAllDomainMappings();
                    LogResult($"🔧 Domain processor loaded with {domainMappings.Count} customer mappings");
                    LogResult("");

                    var weirDomains = domainMappings.Where(d => d.Key.Contains("weir")).ToList();
                    if (weirDomains.Any())
                    {
                        LogResult($"⭐ Found {weirDomains.Count} Weir domain configurations");
                        foreach (var weir in weirDomains)
                        {
                            LogResult($"   🏢 {weir.Key} → {weir.Value?.ToString() ?? "Unknown"}");
                        }
                        LogResult("");
                    }
                }
                catch (Exception domainEx)
                {
                    LogResult($"⚠️ Domain processor warning: {domainEx.Message}");
                    LogResult("");
                }

                // Import emails
                _enhancedProgress?.UpdateProgress(10, 100, "Importing emails and extracting business content...");
                var importSummary = await EmailWorkFlowService.ImportEmailsAsync(_enhancedProgress, LogResult);

                // Process validation for each email with real-time spacing
                if (importSummary?.EmailDetails?.Count > 0)
                {
                    LogResult($"📧 Processing {importSummary.EmailDetails.Count} emails for price validation...");
                    LogResult("");

                    foreach (var email in importSummary.EmailDetails)
                    {
                        var validationResult = ProcessEmailWithPriceValidation(email);
                        var safeEmailDetail = validationResult.CleanedEmail;

                        _enhancedProgress?.IncrementEmailsProcessed();

                        if (_currentRunId != Guid.Empty)
                        {
                            var emailsProcessed = _enhancedProgress?.EmailsProcessed ?? 0;
                            var ordersFound = _enhancedProgress?.OrdersFound ?? 0;
                            var quotesFound = _enhancedProgress?.QuotesFound ?? 0;
                            var revisionsFound = _enhancedProgress?.RevisionsFound ?? 0;

                            _runHistoryManager.UpdateRunProgress(_currentRunId, emailsProcessed, ordersFound,
                                quotesFound, revisionsFound, importSummary.EmailDetails.Count);
                            RefreshRunHistory();
                        }
                        await Task.Delay(1);
                    }
                }

                LogResult($"✅ Email import phase completed. Preparing summary...");
                LogResult("");

                // Display summary (no more complex save/append logic)
                DisplayEmailImportResults(importSummary);

                _enhancedProgress?.UpdateProgress(70, 100, "Starting enhanced MKG injection...");
                await ProcessEnhancedMkgInjection(importSummary);

                // Final completion
                var totalTime = DateTime.Now - _processingStartTime;
                _enhancedProgress?.CompleteOperation($"Enhanced email import completed in {totalTime.TotalSeconds:F1}s!");

                if (_currentRunData != null)
                {
                    PersistCurrentRunDataNow();
                }

                RefreshRunHistory();
                Console.WriteLine($"🎉 Enhanced email import completed in {totalTime.TotalSeconds:F1} seconds");
            }
            catch (Exception ex)
            {
                _isProcessing = false;
                Console.WriteLine($"❌ Enhanced button click error: {ex.Message}");
                LogResult($"❌ Error in enhanced email processing: {ex.Message}");

                _enhancedProgress?.FailOperation(ex.Message);

                if (_currentRunId != Guid.Empty)
                {
                    _runHistoryManager.CompleteRun(_currentRunId, "Failed");
                    RefreshRunHistory();
                }

                UpdateProgress(0, 100, $"Error: {ex.Message}");
                MessageBox.Show($"Error during enhanced email processing: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _isProcessing = false;
                UpdateProgress(-1, -1, "Ready");
            }
        }


        /// <summary>
        /// FINAL WORKAROUND - Completely different method name
        /// </summary>
        public void PersistCurrentRunDataNow()
        {
            try
            {
                if (_currentRunData != null && _currentRunId != Guid.Empty)
                {
                    Console.WriteLine("💾 Saving complete run data with actual content...");

                    // UPDATED: Use renamed methods
                    _currentRunData.EmailResults = CollectEmailResultsForHistory();
                    _currentRunData.MkgResults = CollectMkgResultsForHistory();
                    _currentRunData.ErrorResults = CollectErrorResultsForHistory();

                    // Rest stays the same...
                    _currentRunData.Settings["SavedAt"] = DateTime.Now;
                    _currentRunData.Settings["SavedBy"] = Environment.UserName;
                    _currentRunData.Settings["TotalTabItems"] =
                        (_currentRunData.EmailResults?.Count ?? 0) +
                        (_currentRunData.MkgResults?.Count ?? 0) +
                        (_currentRunData.ErrorResults?.Count ?? 0);

                    _runHistoryManager.SaveRunData(_currentRunData);

                    Console.WriteLine($"✅ Saved complete run data:");
                    Console.WriteLine($"   📧 Email items: {_currentRunData.EmailResults?.Count ?? 0}");
                    Console.WriteLine($"   📊 MKG items: {_currentRunData.MkgResults?.Count ?? 0}");
                    Console.WriteLine($"   ❌ Error items: {_currentRunData.ErrorResults?.Count ?? 0}");
                }
                else
                {
                    Console.WriteLine("⚠️ Cannot save run data - _currentRunData or _currentRunId is null");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error saving complete run data: {ex.Message}");
            }
        }
        public void TestMethod123()
        {
            Console.WriteLine("Test method works!");
        }
        private async void button2_Click(object sender, EventArgs e)
        {
            await ProcessMkgApiTest();
        }

        /// <summary>
        /// Menu MKG Test handler
        /// </summary>
        private async void MkgTestApi_Click(object sender, EventArgs e)
        {
            await ProcessMkgApiTest();
        }
        #endregion

        #region ENHANCED MKG INJECTION LOGIC
        /// <summary>
        /// ENHANCED ProcessMkgInjection - Handles OrderLines + QuoteLines + RevisionLines
        /// ✅ FIXED: Proper counter increments for live statistics
        /// </summary>
        private async Task ProcessEnhancedMkgInjection(EmailImportSummary importSummary)
        {
            if (tabControl != null && tabControl.TabPages.Count > 1)
            {
                tabControl.SelectedIndex = 1; // MKG Results tab
            }
            try
            {
                Console.WriteLine("🚀 Starting ENHANCED MKG injection with real-time updates...");

                // ✅ FIXED: Clear MKG Results tab and prepare for real-time updates
                if (lstMkgResults != null)
                {
                    lstMkgResults.Items.Clear();
                    lstMkgResults.Items.Add("🚀 === LIVE MKG INJECTION MONITOR ===");
                    lstMkgResults.Items.Add($"Started: {DateTime.Now:HH:mm:ss}");
                    lstMkgResults.Items.Add("");
                }

                // Variables to store summaries for failed injections tab
                MkgOrderInjectionSummary orderSummary = null;
                MkgQuoteInjectionSummary quoteSummary = null;
                MkgRevisionInjectionSummary revisionSummary = null;

                // Extract ALL types from email summary using conversion methods
                var allOrderLines = ConvertEmailSummaryToOrderLines(importSummary);
                var allQuoteLines = ConvertEmailSummaryToQuoteLines(importSummary);
                var allRevisionLines = ConvertEmailSummaryToRevisionLines(importSummary);

                // Group by keys
                var orderGroups = allOrderLines.GroupBy(o => o.PoNumber ?? "UNKNOWN").ToList();
                var quoteGroups = allQuoteLines.GroupBy(q => q.RfqNumber ?? "UNKNOWN").ToList();
                var revisionGroups = allRevisionLines.GroupBy(r => r.ArtiCode ?? "UNKNOWN").ToList();

                // Count tracking
                int totalInjected = 0;
                int totalFailed = 0;

                // Show summary
                LogMkgResultRealTime($"📦 Orders to inject: {orderGroups.Count} Orders containing {allOrderLines.Count} OrderLines");
                LogMkgResultRealTime($"💰 Quotes to inject: {quoteGroups.Count} Quotes containing {allQuoteLines.Count} QuoteLines");
                LogMkgResultRealTime($"🔄 Revisions to inject: {revisionGroups.Count} Revision Sets containing {allRevisionLines.Count} RevisionLines");
                LogMkgResultRealTime("");

                // ✅ STEP 1: Orders Injection with Real-time Updates
                if (allOrderLines.Any())
                {
                    LogMkgResultRealTime($"📦 Starting Order injection ({orderGroups.Count} Orders containing {allOrderLines.Count} OrderLines)...");

                    // Show ALL orders - NO TRUNCATION
                    foreach (var orderGroup in orderGroups)
                    {
                        LogMkgResultRealTime($"   📋 PO: {orderGroup.Key} contains {orderGroup.Count()} OrderLines");
                    }
                    LogMkgResultRealTime("");

                    using (var mkgOrderController = new MkgOrderController())
                    {
                        // Use async parallel injection with real-time MKG Results tab updates
                        orderSummary = await mkgOrderController.InjectOrdersAsync(
                            allOrderLines,
                            new Progress<string>(status => LogMkgResultRealTime(status)),
                            _enhancedProgress  // 🎯 ADD THIS LINE!
                        );
                        totalInjected += orderSummary.SuccessfulInjections;
                        totalFailed += orderSummary.FailedInjections;

                        // ✅ FIXED: Increment error counter for each failed injection
                        if (orderSummary.FailedInjections > 0 && _enhancedProgress != null)
                        {
                            for (int i = 0; i < orderSummary.FailedInjections; i++)
                            {
                                _enhancedProgress.IncrementInjectionErrors();
                            }
                        }

                        // Display complete results - NO TRUNCATION
                        DisplayMkgResults(orderSummary);
                    }
                }

                // ✅ STEP 2: Quotes Injection with Real-time Updates
                if (allQuoteLines.Any())
                {
                    LogMkgResultRealTime($"💰 Starting Quote injection ({quoteGroups.Count} Quotes containing {allQuoteLines.Count} QuoteLines)...");

                    // Show ALL quotes - NO TRUNCATION
                    foreach (var quoteGroup in quoteGroups)
                    {
                        LogMkgResultRealTime($"   📋 RFQ: {quoteGroup.Key} contains {quoteGroup.Count()} QuoteLines");
                    }
                    LogMkgResultRealTime("");

                    using (var mkgQuoteController = new MkgQuoteController())
                    {
                        // Use async parallel injection with real-time updates
                        quoteSummary = await mkgQuoteController.InjectQuotesAsync(
                            allQuoteLines,
                            new Progress<string>(status => LogMkgResultRealTime(status))
                        );

                        totalInjected += quoteSummary.SuccessfulInjections;
                        totalFailed += quoteSummary.FailedInjections;

                        // ✅ FIXED: Increment error counter for each failed injection
                        if (quoteSummary.FailedInjections > 0 && _enhancedProgress != null)
                        {
                            for (int i = 0; i < quoteSummary.FailedInjections; i++)
                            {
                                _enhancedProgress.IncrementInjectionErrors();
                            }
                        }

                        // Display complete results - NO TRUNCATION
                        DisplayMkgQuoteResults(quoteSummary);
                    }
                }

                // ✅ STEP 3: Revisions Injection with Real-time Updates
                if (allRevisionLines.Any())
                {
                    LogMkgResultRealTime($"🔄 Starting Revision injection ({revisionGroups.Count} Revision Sets containing {allRevisionLines.Count} RevisionLines)...");

                    // Show ALL revisions - NO TRUNCATION
                    foreach (var revisionGroup in revisionGroups)
                    {
                        LogMkgResultRealTime($"   📋 Article: {revisionGroup.Key} contains {revisionGroup.Count()} RevisionLines");
                    }
                    LogMkgResultRealTime("");

                    using (var mkgRevisionController = new MkgRevisionController())
                    {
                        // Use async parallel injection with real-time updates
                        revisionSummary = await mkgRevisionController.InjectRevisionsAsync(
                            allRevisionLines,
                            new Progress<string>(status => LogMkgResultRealTime(status))
                        );

                        totalInjected += revisionSummary.SuccessfulInjections;
                        totalFailed += revisionSummary.FailedInjections;

                        // ✅ FIXED: Increment error counter for each failed injection
                        if (revisionSummary.FailedInjections > 0 && _enhancedProgress != null)
                        {
                            for (int i = 0; i < revisionSummary.FailedInjections; i++)
                            {
                                _enhancedProgress.IncrementInjectionErrors();
                            }
                        }

                        // Display complete results - NO TRUNCATION
                        DisplayMkgRevisionResults(revisionSummary);
                    }
                }

                // Update Failed Injections Tab with all failure data
                UpdateFailedInjectionsTab(orderSummary, quoteSummary, revisionSummary);

                // Final Summary - Real-time
                LogMkgResultRealTime("");
                LogMkgResultRealTime("=== ENHANCED MKG INJECTION FINISHED ===");
                LogMkgResultRealTime($"📊 INJECTION RESULTS:");
                LogMkgResultRealTime($"   ✅ Total items injected successfully: {totalInjected}");
                LogMkgResultRealTime($"   ❌ Total items failed: {totalFailed}");
                LogMkgResultRealTime($"   📈 Overall success rate: {(totalInjected * 100.0 / Math.Max(totalInjected + totalFailed, 1)):F1}%");

            }
            catch (Exception ex)
            {
                LogMkgResultRealTime($"❌ Critical error in enhanced MKG injection: {ex.Message}");
                Console.WriteLine($"❌ ProcessEnhancedMkgInjection error: {ex.Message}");

                // Increment error counter for the exception
                _enhancedProgress?.IncrementInjectionErrors();
            }
        }
        private void LogMkgResultRealTime(string message)
        {
            try
            {
                if (lstMkgResults != null)
                {
                    if (lstMkgResults.InvokeRequired)
                    {
                        lstMkgResults.Invoke((Action)(() => LogMkgResultRealTime(message)));
                    }
                    else
                    {
                        var timestamp = DateTime.Now.ToString("HH:mm:ss");
                        var logEntry = !string.IsNullOrEmpty(message) ? $"[{timestamp}] {message}" : "";

                        lstMkgResults.Items.Add(logEntry);

                        // Auto-scroll to bottom for real-time effect
                        lstMkgResults.TopIndex = Math.Max(0, lstMkgResults.Items.Count - 1);
                        lstMkgResults.Refresh();
                    }
                }

                // Also log to console
                Console.WriteLine(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in real-time logging: {ex.Message}");
            }
        }

        #endregion

        #region Display Methods

        /// <summary>
        /// MKG Order Results Display - NO TRUNCATION
        /// </summary>
        /// <summary>
        /// Enhanced MKG Results Display - NO TRUNCATION, shows ALL results
        /// </summary>
        /// <summary>
        /// Enhanced MKG Results Display - NO TRUNCATION, shows ALL results with CORRECT properties
        /// </summary>
        /// <summary>
        /// Enhanced MKG Results Display - Fixed with better duplicate handling, success rates, and comprehensive display
        /// </summary>
        private void DisplayMkgResults(
            MkgOrderInjectionSummary orderSummary,
            MkgQuoteInjectionSummary quoteSummary = null,
            MkgRevisionInjectionSummary revisionSummary = null)
        {
            try
            {
                if (lstMkgResults != null)
                {
                    lstMkgResults.Items.Clear();

                    // Header with timestamp
                    lstMkgResults.Items.Add("🚀 === ENHANCED MKG INJECTION RESULTS ===");
                    lstMkgResults.Items.Add($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    lstMkgResults.Items.Add($"User: {Environment.UserName}");
                    lstMkgResults.Items.Add("");

                    // Calculate comprehensive statistics with duplicate tracking
                    var totalItems = (orderSummary?.OrderResults?.Count ?? 0) +
                                   (quoteSummary?.QuoteResults?.Count ?? 0) +
                                   (revisionSummary?.RevisionResults?.Count ?? 0);

                    var totalSuccess = (orderSummary?.SuccessfulInjections ?? 0) +
                                     (quoteSummary?.SuccessfulInjections ?? 0) +
                                     (revisionSummary?.SuccessfulInjections ?? 0);

                    var totalFailed = (orderSummary?.FailedInjections ?? 0) +
                                    (quoteSummary?.FailedInjections ?? 0) +
                                    (revisionSummary?.FailedInjections ?? 0);

                    // Calculate duplicates separately
                    var orderDuplicates = orderSummary?.OrderResults?.Count(r => r.ErrorMessage?.Contains("Duplicate") == true) ?? 0;
                    var quoteDuplicates = quoteSummary?.QuoteResults?.Count(r => r.ErrorMessage?.Contains("Duplicate") == true) ?? 0;
                    var revisionDuplicates = revisionSummary?.RevisionResults?.Count(r => r.ErrorMessage?.Contains("Duplicate") == true) ?? 0;
                    var totalDuplicates = orderDuplicates + quoteDuplicates + revisionDuplicates;

                    var realFailures = totalFailed - totalDuplicates;
                    var actualSuccessRate = totalItems > 0 ? ((totalSuccess + totalDuplicates) * 100.0 / totalItems) : 0;

                    // Enhanced Summary Statistics
                    lstMkgResults.Items.Add("📊 === INJECTION SUMMARY ===");
                    lstMkgResults.Items.Add($"📦 Total Items Processed: {totalItems}");
                    lstMkgResults.Items.Add($"✅ Successfully Injected: {totalSuccess}");
                    lstMkgResults.Items.Add($"🔄 Duplicates Skipped: {totalDuplicates}");
                    lstMkgResults.Items.Add($"❌ Real Failures: {realFailures}");
                    lstMkgResults.Items.Add($"📈 Effective Success Rate: {actualSuccessRate:F1}% (including duplicates as handled)");
                    lstMkgResults.Items.Add("");

                    // Orders Section - Enhanced with better duplicate handling
                    if (orderSummary != null && orderSummary.OrderResults.Any())
                    {
                        lstMkgResults.Items.Add("📦 === ORDER INJECTION RESULTS ===");
                        lstMkgResults.Items.Add($"📊 Headers Created: {orderSummary.TotalOrders}");
                        lstMkgResults.Items.Add($"📋 Lines Processed: {orderSummary.OrderResults.Count}");
                        lstMkgResults.Items.Add($"✅ Successful: {orderSummary.SuccessfulInjections}");
                        lstMkgResults.Items.Add($"🔄 Duplicates: {orderDuplicates}");
                        lstMkgResults.Items.Add($"❌ Real Failures: {orderSummary.FailedInjections - orderDuplicates}");
                        lstMkgResults.Items.Add($"⏱️ Processing Time: {orderSummary.ProcessingTime.TotalSeconds:F1}s");
                        lstMkgResults.Items.Add("");

                        // Group by PO Number with enhanced display
                        var orderGroups = orderSummary.OrderResults
                            .GroupBy(r => r.PoNumber)
                            .ToList();

                        if (orderGroups.Any())
                        {
                            lstMkgResults.Items.Add($"📦 Order Details ({orderGroups.Count} unique POs):");
                            foreach (var group in orderGroups)
                            {
                                var lines = group.ToList();
                                var successCount = lines.Count(l => l.Success);
                                var duplicateCount = lines.Count(l => l.ErrorMessage?.Contains("Duplicate") == true);
                                var realFailureCount = lines.Count(l => !l.Success && !(l.ErrorMessage?.Contains("Duplicate") == true));
                                var poNumber = group.Key ?? "UNKNOWN";

                                lstMkgResults.Items.Add($"   ├── PO: {poNumber} (✅{successCount} 🔄{duplicateCount} ❌{realFailureCount})");

                                foreach (var line in lines)
                                {
                                    string status, details;
                                    if (line.Success)
                                    {
                                        status = "✅";
                                        details = $"{line.ArtiCode}";
                                        // Add description if available and not empty
                                        if (!string.IsNullOrEmpty(line.Description))
                                        {
                                            details += $" | {line.Description}";
                                        }
                                    }
                                    else if (line.ErrorMessage?.Contains("Duplicate") == true)
                                    {
                                        status = "🔄";
                                        details = $"{line.ArtiCode} | DUPLICATE";
                                    }
                                    else
                                    {
                                        status = "❌";
                                        details = $"{line.ArtiCode} | ERROR: {line.ErrorMessage}";
                                    }

                                    lstMkgResults.Items.Add($"   │   {status} {details}");
                                }
                            }
                        }
                        lstMkgResults.Items.Add("");
                    }

                    // Quotes Section - Enhanced display
                    if (quoteSummary != null && quoteSummary.QuoteResults.Any())
                    {
                        lstMkgResults.Items.Add("💰 === QUOTE INJECTION RESULTS ===");
                        lstMkgResults.Items.Add($"📊 Headers Created: {quoteSummary.TotalQuotes}");
                        lstMkgResults.Items.Add($"📋 Lines Processed: {quoteSummary.QuoteResults.Count}");
                        lstMkgResults.Items.Add($"✅ Successful: {quoteSummary.SuccessfulInjections}");
                        lstMkgResults.Items.Add($"🔄 Duplicates: {quoteDuplicates}");
                        lstMkgResults.Items.Add($"❌ Real Failures: {quoteSummary.FailedInjections - quoteDuplicates}");
                        lstMkgResults.Items.Add($"⏱️ Processing Time: {quoteSummary.ProcessingTime.TotalSeconds:F1}s");
                        lstMkgResults.Items.Add("");

                        // Group by RFQ Number
                        var quoteGroups = quoteSummary.QuoteResults
                            .GroupBy(r => r.RfqNumber)
                            .ToList();

                        if (quoteGroups.Any())
                        {
                            lstMkgResults.Items.Add($"💰 Quote Details ({quoteGroups.Count} unique RFQs):");
                            foreach (var group in quoteGroups)
                            {
                                var lines = group.ToList();
                                var successCount = lines.Count(l => l.Success);
                                var duplicateCount = lines.Count(l => l.ErrorMessage?.Contains("Duplicate") == true);
                                var realFailureCount = lines.Count(l => !l.Success && !(l.ErrorMessage?.Contains("Duplicate") == true));
                                var rfqNumber = group.Key ?? "UNKNOWN";

                                lstMkgResults.Items.Add($"   ├── RFQ: {rfqNumber} (✅{successCount} 🔄{duplicateCount} ❌{realFailureCount})");

                                foreach (var line in lines)
                                {
                                    string status, details;
                                    if (line.Success)
                                    {
                                        status = "✅";
                                        details = $"{line.ArtiCode}";
                                        if (!string.IsNullOrEmpty(line.QuotedPrice))
                                        {
                                            details += $" | Price: {line.QuotedPrice}";
                                        }
                                    }
                                    else if (line.ErrorMessage?.Contains("Duplicate") == true)
                                    {
                                        status = "🔄";
                                        details = $"{line.ArtiCode} | DUPLICATE";
                                    }
                                    else
                                    {
                                        status = "❌";
                                        details = $"{line.ArtiCode} | ERROR: {line.ErrorMessage}";
                                    }

                                    lstMkgResults.Items.Add($"   │   {status} {details}");
                                }
                            }
                        }
                        lstMkgResults.Items.Add("");
                    }

                    // Revisions Section - Enhanced display
                    if (revisionSummary != null && revisionSummary.RevisionResults.Any())
                    {
                        lstMkgResults.Items.Add("🔄 === REVISION INJECTION RESULTS ===");
                        lstMkgResults.Items.Add($"📊 Headers Created: {revisionSummary.TotalRevisions}");
                        lstMkgResults.Items.Add($"📋 Lines Processed: {revisionSummary.RevisionResults.Count}");
                        lstMkgResults.Items.Add($"✅ Successful: {revisionSummary.SuccessfulInjections}");
                        lstMkgResults.Items.Add($"🔄 Duplicates: {revisionDuplicates}");
                        lstMkgResults.Items.Add($"❌ Real Failures: {revisionSummary.FailedInjections - revisionDuplicates}");
                        lstMkgResults.Items.Add($"⏱️ Processing Time: {revisionSummary.ProcessingTime.TotalSeconds:F1}s");
                        lstMkgResults.Items.Add("");

                        // Group by Article Code
                        var revisionGroups = revisionSummary.RevisionResults
                            .GroupBy(r => r.ArtiCode)
                            .ToList();

                        if (revisionGroups.Any())
                        {
                            lstMkgResults.Items.Add($"🔄 Revision Details ({revisionGroups.Count} unique articles):");
                            foreach (var group in revisionGroups)
                            {
                                var lines = group.ToList();
                                var successCount = lines.Count(l => l.Success);
                                var duplicateCount = lines.Count(l => l.ErrorMessage?.Contains("Duplicate") == true);
                                var realFailureCount = lines.Count(l => !l.Success && !(l.ErrorMessage?.Contains("Duplicate") == true));
                                var articleCode = group.Key ?? "UNKNOWN";

                                lstMkgResults.Items.Add($"   ├── Article: {articleCode} (✅{successCount} 🔄{duplicateCount} ❌{realFailureCount})");

                                foreach (var line in lines)
                                {
                                    string status, details;
                                    if (line.Success)
                                    {
                                        status = "✅";
                                        details = $"{line.CurrentRevision}→{line.NewRevision}";
                                    }
                                    else if (line.ErrorMessage?.Contains("Duplicate") == true)
                                    {
                                        status = "🔄";
                                        details = $"{line.ArtiCode} | DUPLICATE";
                                    }
                                    else
                                    {
                                        status = "❌";
                                        details = $"{line.ArtiCode} | ERROR: {line.ErrorMessage}";
                                    }

                                    lstMkgResults.Items.Add($"   │   {status} {details}");
                                }
                            }
                        }
                        lstMkgResults.Items.Add("");
                    }

                    // 🎯 INCREMENTAL PROCESSING STEPS - Fixed Integration
                    var mkgProcessingLog = EmailWorkFlowService.GetMkgProcessingLog();
                    if (mkgProcessingLog.Any())
                    {
                        lstMkgResults.Items.Add("🔄 === INCREMENTAL PROCESSING STEPS ===");
                        foreach (var logEntry in mkgProcessingLog)
                        {
                            lstMkgResults.Items.Add($"   {logEntry}");
                        }
                        lstMkgResults.Items.Add("");
                    }

                    lstMkgResults.Items.Add("=== END OF MKG INJECTION RESULTS ===");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error displaying MKG results: {ex.Message}");
                if (lstMkgResults != null)
                {
                    lstMkgResults.Items.Add($"❌ Error displaying results: {ex.Message}");
                }
            }
        }
        /// <summary>
        /// Quote Results Display - NO TRUNCATION
        /// </summary>
        private void DisplayMkgQuoteResults(MkgQuoteInjectionSummary summary)
        {
            try
            {
                if (summary == null) return;

                if (lstMkgResults != null)
                {
                    lstMkgResults.Items.Add($"=== QUOTE INJECTION RESULTS ===");
                    lstMkgResults.Items.Add($"Headers Created: {summary.TotalQuotes} | Lines Processed: {summary.QuoteResults.Count}");
                    lstMkgResults.Items.Add($"Successful: {summary.SuccessfulInjections} | Failed: {summary.FailedInjections}");
                    lstMkgResults.Items.Add($"Processing Time: {summary.ProcessingTime.TotalSeconds:F1}s");
                    lstMkgResults.Items.Add("");

                    // Show ALL Quote Headers
                    var quoteGroups = summary.QuoteResults
                        .Where(r => r.Success && !string.IsNullOrEmpty(r.MkgQuoteId))
                        .GroupBy(r => r.MkgQuoteId)
                        .ToList();

                    lstMkgResults.Items.Add($"Showing ALL {quoteGroups.Count} Quote Headers:");
                    lstMkgResults.Items.Add("");

                    foreach (var quoteGroup in quoteGroups)
                    {
                        var headerId = quoteGroup.Key;
                        var quoteLines = quoteGroup.ToList();
                        var firstLine = quoteLines.First();
                        var lineCount = quoteLines.Count;

                        lstMkgResults.Items.Add($"├── PARENT: Quote Header {headerId}");
                        lstMkgResults.Items.Add($"│   ├── RFQ: {firstLine.RfqNumber} | Lines: {lineCount} | Time: {firstLine.ProcessedAt:HH:mm:ss}");
                        lstMkgResults.Items.Add($"│   └── API: vofh table (Primary Key: vofh_num={headerId})");

                        // Show ALL quote lines
                        foreach (var line in quoteLines)
                        {
                            var status = line.Success ? "✓" : "✗";
                            var price = !string.IsNullOrEmpty(line.QuotedPrice) ? $"(€{line.QuotedPrice})" : "(€)";
                            lstMkgResults.Items.Add($"│       └── CHILD: {line.ArtiCode} {price} → {status}");
                        }
                        lstMkgResults.Items.Add("│");
                    }
                    lstMkgResults.Items.Add("└── END OF QUOTE HEADERS");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error displaying MKG quote results: {ex.Message}");
            }
        }

        /// <summary>
        /// Revision Results Display - NO TRUNCATION
        /// </summary>
        private void DisplayMkgRevisionResults(MkgRevisionInjectionSummary summary)
        {
            try
            {
                if (summary == null) return;

                if (lstMkgResults != null)
                {
                    lstMkgResults.Items.Add($"=== REVISION INJECTION RESULTS ===");
                    lstMkgResults.Items.Add($"Headers Created: {summary.TotalRevisions} | Changes Processed: {summary.RevisionResults.Count}");
                    lstMkgResults.Items.Add($"Successful: {summary.SuccessfulInjections} | Failed: {summary.FailedInjections}");
                    lstMkgResults.Items.Add($"Processing Time: {summary.ProcessingTime.TotalSeconds:F1}s");
                    lstMkgResults.Items.Add("");

                    lstMkgResults.Items.Add("REVISION INJECTION WORKFLOW:");
                    lstMkgResults.Items.Add("   1. Group RevisionLines by Article Code");
                    lstMkgResults.Items.Add("   2. CREATE REVISION HEADER: POST /Documents/revision_header → get revision_id");
                    lstMkgResults.Items.Add("   3. INJECT REVISION LINES: POST /Documents/revision_lines (refs revision_id)");
                    lstMkgResults.Items.Add("");

                    // Show ALL Revision Headers
                    var revisionGroups = summary.RevisionResults
                        .Where(r => r.Success && !string.IsNullOrEmpty(r.MkgRevisionId))
                        .GroupBy(r => r.MkgRevisionId)
                        .ToList();

                    lstMkgResults.Items.Add($"COMPLETE PARENT → CHILD STRUCTURE:");
                    lstMkgResults.Items.Add($"Showing ALL {revisionGroups.Count} Revision Headers:");
                    lstMkgResults.Items.Add("");

                    foreach (var revisionGroup in revisionGroups)
                    {
                        var headerId = revisionGroup.Key;
                        var revisionLines = revisionGroup.ToList();
                        var firstLine = revisionLines.First();
                        var lineCount = revisionLines.Count;

                        lstMkgResults.Items.Add($"├── PARENT: Revision Header {headerId}");
                        lstMkgResults.Items.Add($"│   ├── Article: {firstLine.ArtiCode} | Changes: {lineCount} | Time: {firstLine.ProcessedAt:HH:mm:ss}");
                        lstMkgResults.Items.Add($"│   └── API: revision_header table (Primary Key: revision_id={headerId})");

                        // Show ALL revision lines
                        foreach (var line in revisionLines)
                        {
                            var status = line.Success ? "✓" : "✗";
                            var change = $"('{line.OldValue}' → '{line.NewValue}')";
                            lstMkgResults.Items.Add($"│       └── CHILD: Revision {change} → {status}");
                            lstMkgResults.Items.Add($"│           ├── Reason: {line.ChangeReason}");
                        }
                        lstMkgResults.Items.Add("│");
                    }
                    lstMkgResults.Items.Add("└── END OF REVISION HEADERS");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error displaying MKG revision results: {ex.Message}");
            }
        }

        #endregion

        #region Failed Injections Display

        /// <summary>
        /// Complete Failed Injections Display for Orders, Quotes, and Revisions
        /// </summary>
        // Replace your DisplayAllFailedInjections method with this fixed version:
        private void DisplayAllFailedInjections(MkgOrderInjectionSummary orderSummary,
                                MkgQuoteInjectionSummary quoteSummary,
                                MkgRevisionInjectionSummary revisionSummary)
        {
            try
            {
                if (lstFailedInjections == null) return;

                lstFailedInjections.Items.Clear();

                // 🔧 FIX: Calculate REAL failures vs duplicates separately
                int totalRealFailures = 0;
                int totalMkgDuplicates = 0;

                if (orderSummary != null)
                {
                    // Count only NON-duplicate failures as real failures
                    var realOrderFailures = orderSummary.OrderResults.Count(r =>
                        !r.Success && !r.HttpStatusCode.Contains("DUPLICATE"));
                    totalRealFailures += realOrderFailures;

                    // Count MKG duplicates separately
                    totalMkgDuplicates += orderSummary.OrderResults.Count(r =>
                        r.HttpStatusCode.Contains("DUPLICATE"));
                }
                if (quoteSummary != null)
                {
                    var realQuoteFailures = quoteSummary.QuoteResults.Count(r =>
                        !r.Success && !r.HttpStatusCode.Contains("DUPLICATE"));
                    totalRealFailures += realQuoteFailures;

                    totalMkgDuplicates += quoteSummary.QuoteResults.Count(r =>
                        r.HttpStatusCode.Contains("DUPLICATE"));
                }
                if (revisionSummary != null)
                {
                    var realRevisionFailures = revisionSummary.RevisionResults.Count(r =>
                        !r.Success && !r.HttpStatusCode.Contains("DUPLICATE"));
                    totalRealFailures += realRevisionFailures;

                    totalMkgDuplicates += revisionSummary.RevisionResults.Count(r =>
                        r.HttpStatusCode.Contains("DUPLICATE"));
                }

                // 🔧 CRITICAL FIX: Set flags correctly
                _hasInjectionFailures = (totalRealFailures > 0);  // Only true for REAL failures
                _hasMkgDuplicatesDetected = (totalMkgDuplicates > 0);  // True for duplicates

                Console.WriteLine($"🔧 FIXED FLAGS: _hasInjectionFailures={_hasInjectionFailures}, _hasMkgDuplicatesDetected={_hasMkgDuplicatesDetected}");
                Console.WriteLine($"🔧 COUNTS: Real failures={totalRealFailures}, Duplicates={totalMkgDuplicates}");

                // 🔧 FIX: Update current run data correctly
                if (_currentRunData != null)
                {
                    _currentRunData.HasInjectionFailures = (totalRealFailures > 0);
                    _currentRunData.TotalFailuresAtCompletion = totalRealFailures;
                    _currentRunData.HasMkgDuplicatesDetected = (totalMkgDuplicates > 0);
                    _currentRunData.TotalDuplicatesDetected = totalMkgDuplicates;
                }

                // Update status label
                if (lblFailedStatus != null)
                {
                    if (totalRealFailures == 0 && totalMkgDuplicates == 0)
                    {
                        lblFailedStatus.Text = "No issues detected - all injections successful!";
                        lblFailedStatus.ForeColor = Color.Green;
                    }
                    else if (totalRealFailures > 0)
                    {
                        lblFailedStatus.Text = $"{totalRealFailures} failures + {totalMkgDuplicates} MKG duplicates detected";
                        lblFailedStatus.ForeColor = Color.Red;
                    }
                    else // Only MKG duplicates, no real failures
                    {
                        lblFailedStatus.Text = $"{totalMkgDuplicates} MKG duplicates detected (skipped automatically)";
                        lblFailedStatus.ForeColor = Color.DarkOrange;
                    }
                }

                // 🔧 FIX: Update tab status with correct values
                UpdateTabStatus();

                // Header
                lstFailedInjections.Items.Add("=== ENHANCED ERROR TRACKING ===");
                lstFailedInjections.Items.Add("Intelligent error analysis enabled");
                lstFailedInjections.Items.Add("");

                if (totalRealFailures == 0 && totalMkgDuplicates == 0)
                {
                    lstFailedInjections.Items.Add("✅ NO INJECTION FAILURES!");
                    lstFailedInjections.Items.Add("✅ All injections completed successfully");
                    lstFailedInjections.Items.Add("✅ Ready for production use");
                    return;
                }

                // Show MKG duplicates first (these are good)
                if (totalMkgDuplicates > 0)
                {
                    lstFailedInjections.Items.Add($"🔄 MKG DUPLICATES: {totalMkgDuplicates} items");
                    lstFailedInjections.Items.Add("ℹ️ These are EXPECTED and indicate correct duplicate prevention");
                    lstFailedInjections.Items.Add("ℹ️ Items already exist in MKG - no action needed");
                    lstFailedInjections.Items.Add("");
                }

                // Show actual injection failures after duplicates
                if (totalRealFailures > 0)
                {
                    lstFailedInjections.Items.Add($"❌ REAL INJECTION FAILURES: {totalRealFailures} total");
                    lstFailedInjections.Items.Add("");

                    // Display order failures
                    if (orderSummary != null)
                        DisplayOrderFailures(orderSummary);

                    // Display quote failures  
                    if (quoteSummary != null)
                        DisplayQuoteFailures(quoteSummary);

                    // Display revision failures
                    if (revisionSummary != null)
                        DisplayRevisionFailures(revisionSummary);
                }

                // Summary at the bottom
                lstFailedInjections.Items.Add("");
                lstFailedInjections.Items.Add("=== SUMMARY ===");
                if (totalMkgDuplicates > 0)
                    lstFailedInjections.Items.Add($"🔄 MKG duplicates detected: {totalMkgDuplicates}");
                if (totalRealFailures > 0)
                    lstFailedInjections.Items.Add($"❌ Failed injections: {totalRealFailures}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error displaying failed injections: {ex.Message}");
            }
        }
       
        private void DisplayDuplicateStatistics(MkgOrderInjectionSummary orderSummary,
                                       MkgQuoteInjectionSummary quoteSummary,
                                       MkgRevisionInjectionSummary revisionSummary)
        {
            var totalDuplicates = 0;

            if (orderSummary != null)
                totalDuplicates += orderSummary.OrderResults.Count(r => r.HttpStatusCode.Contains("DUPLICATE"));
            if (quoteSummary != null)
                totalDuplicates += quoteSummary.QuoteResults.Count(r => r.HttpStatusCode.Contains("DUPLICATE"));
            if (revisionSummary != null)
                totalDuplicates += revisionSummary.RevisionResults.Count(r => r.HttpStatusCode.Contains("DUPLICATE"));

            if (totalDuplicates > 0)
            {
                lstFailedInjections.Items.Add("");
                lstFailedInjections.Items.Add($"🔄 Duplicates detected: {totalDuplicates}");
                lstFailedInjections.Items.Add("");
                lstFailedInjections.Items.Add("These duplicates were automatically handled:");
                lstFailedInjections.Items.Add("• Cross-email duplicates were filtered out during processing");
                lstFailedInjections.Items.Add("• MKG system duplicates were detected and skipped");
                lstFailedInjections.Items.Add("• No manual action required - duplicates prevent data corruption");
            }
        }
        #endregion

        #region MKG API Test Logic
        private async Task ProcessMkgApiTest()
        {
            try
            {
                Console.WriteLine("=== MKG TEST STARTING ===");

                // ✅ FIX: Switch to MKG Results tab when API test starts
                if (tabControl != null && tabControl.TabPages.Count > 1)
                {
                    tabControl.SelectedIndex = 1; // Switch to MKG Results tab (index 1)
                }

                if (lstMkgResults != null)
                {
                    lstMkgResults.Items.Clear();
                    lstMkgResults.Items.Add("Initializing enhanced MKG test suite...");
                    lstMkgResults.Update();
                }

                toolStripStatusLabel.Text = "Running enhanced MKG tests...";
                ShowProgress("Starting MKG tests...", true, 100, 0);

                using (var tester = new Debug.MkgApiTester())
                {
                    tester.TestOrderCount = int.Parse(ConfigurationManager.AppSettings["MkgTest:DefaultOrderCount"] ?? "2");
                    tester.TestQuoteCount = int.Parse(ConfigurationManager.AppSettings["MkgTest:DefaultQuoteCount"] ?? "2");
                    tester.TestRevisionCount = int.Parse(ConfigurationManager.AppSettings["MkgTest:DefaultRevisionCount"] ?? "2");

                    var progress = new Progress<(int current, int total, string status)>((p) =>
                    {
                        try
                        {
                            UpdateProgress(p.current * 100 / Math.Max(p.total, 1), 100, p.status);
                            if (lstMkgResults != null)
                            {
                                lstMkgResults.Items.Add($"[{p.current}/{p.total}] {p.status}");
                                lstMkgResults.TopIndex = lstMkgResults.Items.Count - 1;
                                lstMkgResults.Update();
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Progress update error: {ex.Message}");
                        }
                    });

                    var results = await tester.RunCompleteTestAsync(progress);
                    HideProgress();

                    if (results != null)
                    {
                        DisplayMkgTestResults(results);
                        PopulateFailedInjectionsTab(results);

                        if (results.AllTestsPassed)
                        {
                            toolStripStatusLabel.Text = $"All tests passed! ({results.TotalTestTime.TotalSeconds:F1}s)";
                        }
                        else
                        {
                            var workingCount = (results.OrdersWorking ? 1 : 0) + (results.QuotesWorking ? 1 : 0) + (results.RevisionsWorking ? 1 : 0);
                            toolStripStatusLabel.Text = $"{workingCount}/3 systems working ({results.TotalTestTime.TotalSeconds:F1}s)";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in ProcessMkgApiTest: {ex.Message}");
                HideProgress();
                toolStripStatusLabel.Text = "MKG test failed";
                if (lstMkgResults != null)
                {
                    lstMkgResults.Items.Add($"❌ Test error: {ex.Message}");
                }
            }
        }
        private void DisplayMkgTestResults(MkgTestResults results)
        {
            try
            {
                if (lstMkgResults != null)
                {
                    lstMkgResults.Items.Clear();
                    lstMkgResults.Items.Add("=== MKG TEST RESULTS ===");
                    lstMkgResults.Items.Add($"Test Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    lstMkgResults.Items.Add("");

                    if (!string.IsNullOrEmpty(results.FullReport))
                    {
                        var reportLines = results.FullReport.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var line in reportLines)
                        {
                            lstMkgResults.Items.Add(line);
                        }
                    }
                    else
                    {
                        lstMkgResults.Items.Add($"Configuration Valid: {results.ConfigurationValid}");
                        lstMkgResults.Items.Add($"API Login Success: {results.ApiLoginSuccess}");
                        lstMkgResults.Items.Add($"Orders Working: {results.OrdersWorking} ({results.OrdersProcessed} processed)");
                        lstMkgResults.Items.Add($"Quotes Working: {results.QuotesWorking} ({results.QuotesProcessed} processed)");
                        lstMkgResults.Items.Add($"Revisions Working: {results.RevisionsWorking} ({results.RevisionsProcessed} processed)");
                        lstMkgResults.Items.Add($"Test Duration: {results.TotalTestTime.TotalSeconds:F1}s");
                    }

                    lstMkgResults.TopIndex = 0;
                    lstMkgResults.Refresh();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error displaying MKG test results: {ex.Message}");
            }
        }

        private void PopulateFailedInjectionsTab(MkgTestResults results)
        {
            try
            {
                if (lstFailedInjections == null || results == null)
                    return;

                lstFailedInjections.Items.Clear();

                var hasFailures = !results.ConfigurationValid || !results.ApiLoginSuccess ||
                                 !results.OrdersWorking || !results.QuotesWorking || !results.RevisionsWorking;

                if (!hasFailures)
                {
                    if (lblFailedStatus != null)
                    {
                        lblFailedStatus.Text = "No failures detected - all systems working!";
                        lblFailedStatus.ForeColor = Color.Green;
                    }

                    lstFailedInjections.Items.Add("NO FAILURES DETECTED!");
                    lstFailedInjections.Items.Add("");
                    lstFailedInjections.Items.Add("All systems are working correctly:");
                    lstFailedInjections.Items.Add("   Configuration is valid");
                    lstFailedInjections.Items.Add("   API connection successful");
                    lstFailedInjections.Items.Add("   Order system working");
                    lstFailedInjections.Items.Add("   Quote system working");
                    lstFailedInjections.Items.Add("   Revision system working");
                    lstFailedInjections.Items.Add("");
                    lstFailedInjections.Items.Add($"Test completed in {results.TotalTestTime.TotalSeconds:F1}s");
                    lstFailedInjections.Items.Add($"Systems working: 3/3");
                    lstFailedInjections.Items.Add($"Orders processed: {results.OrdersProcessed}");
                    lstFailedInjections.Items.Add($"Quotes processed: {results.QuotesProcessed}");
                    lstFailedInjections.Items.Add($"Revisions processed: {results.RevisionsProcessed}");
                    return;
                }

                if (lblFailedStatus != null)
                {
                    var workingSystemsCount = (results.OrdersWorking ? 1 : 0) + (results.QuotesWorking ? 1 : 0) + (results.RevisionsWorking ? 1 : 0);
                    lblFailedStatus.Text = $"{workingSystemsCount}/3 systems working - see details below";
                    lblFailedStatus.ForeColor = Color.Red;
                }

                lstFailedInjections.Items.Add("=== FAILED INJECTIONS & ISSUES ===");
                lstFailedInjections.Items.Add($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                lstFailedInjections.Items.Add("");

                if (!results.ConfigurationValid)
                {
                    lstFailedInjections.Items.Add("CONFIGURATION ISSUES:");
                    lstFailedInjections.Items.Add("   • Check App.config for missing MKG API settings");
                    lstFailedInjections.Items.Add("   • Verify MkgApi:Urls:Base, Username, Password, ApiKey");
                    lstFailedInjections.Items.Add("");
                }

                if (!results.ApiLoginSuccess)
                {
                    lstFailedInjections.Items.Add("API CONNECTION FAILURES:");
                    lstFailedInjections.Items.Add("   • MKG API login failed");
                    lstFailedInjections.Items.Add("   • Check network connectivity to MKG server");
                    lstFailedInjections.Items.Add("   • Verify credentials are correct");
                    lstFailedInjections.Items.Add("");
                }

                if (!results.OrdersWorking)
                {
                    lstFailedInjections.Items.Add("ORDER SYSTEM FAILURES:");
                    lstFailedInjections.Items.Add($"   • Processed: {results.OrdersProcessed} orders");
                    lstFailedInjections.Items.Add("   • Check MkgOrderController functionality");
                    lstFailedInjections.Items.Add("   • Verify order injection logic");
                    lstFailedInjections.Items.Add("");
                }

                if (!results.QuotesWorking)
                {
                    lstFailedInjections.Items.Add("QUOTE SYSTEM FAILURES:");
                    lstFailedInjections.Items.Add($"   • Processed: {results.QuotesProcessed} quotes");
                    lstFailedInjections.Items.Add("   • Check MkgQuoteController functionality");
                    lstFailedInjections.Items.Add("   • Verify quote injection logic");
                    lstFailedInjections.Items.Add("");
                }

                if (!results.RevisionsWorking)
                {
                    lstFailedInjections.Items.Add("REVISION SYSTEM FAILURES:");
                    lstFailedInjections.Items.Add($"   • Processed: {results.RevisionsProcessed} revisions");
                    lstFailedInjections.Items.Add("   • Check MkgRevisionController functionality");
                    lstFailedInjections.Items.Add("   • Verify revision injection logic");
                    lstFailedInjections.Items.Add("");
                }

                var workingSystemsCount2 = (results.OrdersWorking ? 1 : 0) + (results.QuotesWorking ? 1 : 0) + (results.RevisionsWorking ? 1 : 0);
                lstFailedInjections.Items.Add("=== SUMMARY ===");
                lstFailedInjections.Items.Add($"{workingSystemsCount2}/3 injection systems working");
                lstFailedInjections.Items.Add($"Total test time: {results.TotalTestTime.TotalSeconds:F1}s");

                lstFailedInjections.Items.Add("");
                lstFailedInjections.Items.Add("TROUBLESHOOTING RECOMMENDATIONS:");
                lstFailedInjections.Items.Add("   • Check MKG API connectivity and credentials");
                lstFailedInjections.Items.Add("   • Verify network connection to MKG server");
                lstFailedInjections.Items.Add("   • Review configuration settings in App.config");
                lstFailedInjections.Items.Add("   • Contact MKG system administrator if issues persist");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR in PopulateFailedInjectionsTab: {ex.Message}");
                if (lstFailedInjections != null)
                {
                    lstFailedInjections.Items.Add($"Error populating failed injections: {ex.Message}");
                }
            }
        }
        #endregion

        #region Menu Setup
        private void SetupMkgMenuStrip()
        {
            try
            {
                MkgMenuButton.DropDownItems.Clear();

                ToolStripMenuItem mkgEmailImportMenuItem = new ToolStripMenuItem("MKG Email Import");
                mkgEmailImportMenuItem.Click += (sender, e) => Email_Import_Tick(sender, e);
                MkgMenuButton.DropDownItems.Add(mkgEmailImportMenuItem);

                ToolStripMenuItem mkgTestApiMenuItem = new ToolStripMenuItem("Mkg Test API");
                mkgTestApiMenuItem.Click += (sender, e) => button2_Click(sender, e);
                MkgMenuButton.DropDownItems.Add(mkgTestApiMenuItem);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error setting up MKG menu: {ex.Message}");
            }
        }

        private void SetupBottomToolStrip()
        {
            try
            {
                bottomToolStrip = new ToolStrip();
                bottomToolStrip.Dock = DockStyle.Bottom;
                bottomToolStrip.GripStyle = ToolStripGripStyle.Hidden;
                bottomToolStrip.BackColor = SystemColors.Control;
                bottomToolStrip.Height = 25;

                toolStripStatusLabel = new ToolStripStatusLabel();
                toolStripStatusLabel.Text = "Ready";
                toolStripStatusLabel.Spring = true;
                toolStripStatusLabel.TextAlign = ContentAlignment.MiddleLeft;

                toolStripProgressLabel = new ToolStripLabel();
                toolStripProgressLabel.Text = "";
                toolStripProgressLabel.Visible = false;
                toolStripProgressLabel.Margin = new Padding(10, 0, 10, 0);

                toolStripProgressBar = new ToolStripProgressBar();
                toolStripProgressBar.Size = new Size(200, 16);
                toolStripProgressBar.Visible = false;
                toolStripProgressBar.Style = ProgressBarStyle.Continuous;

                bottomToolStrip.Items.Add(toolStripStatusLabel);
                bottomToolStrip.Items.Add(toolStripProgressLabel);
                bottomToolStrip.Items.Add(toolStripProgressBar);

                this.Controls.Add(bottomToolStrip);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error setting up bottom toolbar: {ex.Message}");
            }
        }
        #endregion

        #region Tab Setup Methods
        private void InitializeAllTabControls()
        {
            try
            {
                SetupEmailImportTab();
                SetupMkgResultsTab();
                SetupFailedInjectionsTab();
                this.MinimumSize = new Size(900, 600);
                Console.WriteLine("All tab controls initialized successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error initializing tab controls: {ex.Message}");
                MessageBox.Show($"Error initializing tab controls: {ex.Message}", "Initialization Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Utility Methods
        private void UpdateProgress(int current, int total, string status = null)
        {
            try
            {
                if (!string.IsNullOrEmpty(status))
                {
                    toolStripStatusLabel.Text = status;
                }

                if (current < 0 || total <= 0)
                {
                    toolStripProgressBar.Visible = false;
                    toolStripProgressLabel.Visible = false;
                    return;
                }

                current = Math.Max(0, Math.Min(current, total));

                toolStripProgressBar.Maximum = total;
                toolStripProgressBar.Value = current;
                toolStripProgressBar.Visible = true;
                toolStripProgressLabel.Text = $"{current} / {total}";
                toolStripProgressLabel.Visible = true;

                Application.DoEvents();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Progress bar error: {ex.Message}");
                toolStripProgressBar.Visible = false;
                toolStripProgressLabel.Visible = false;
            }
        }

        private void ShowProgress(string status, bool showProgressBar = true, int maximum = 100, int current = 0)
        {
            try
            {
                toolStripStatusLabel.Text = status;

                if (showProgressBar && maximum > 0 && current >= 0)
                {
                    toolStripProgressBar.Maximum = maximum;
                    toolStripProgressBar.Value = Math.Min(current, maximum);
                    toolStripProgressBar.Visible = true;
                    toolStripProgressLabel.Visible = true;
                }
                else
                {
                    toolStripProgressBar.Visible = false;
                    toolStripProgressLabel.Visible = false;
                }

                Application.DoEvents();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ShowProgress error: {ex.Message}");
                toolStripProgressBar.Visible = false;
                toolStripProgressLabel.Visible = false;
            }
        }

        private void HideProgress()
        {
            try
            {
                toolStripProgressBar.Visible = false;
                toolStripProgressLabel.Visible = false;
                toolStripStatusLabel.Text = "Ready";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error hiding progress: {ex.Message}");
            }
        }

        private void LogResult(string message)
        {
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action<string>(LogResult), message);
                    return;
                }

                if (lstEmailImportResults != null)
                {
                    var timestamp = DateTime.Now.ToString("HH:mm:ss");
                    lstEmailImportResults.Items.Add($"[{timestamp}] {message}");

                    if (lstEmailImportResults.Items.Count > 0)
                    {
                        lstEmailImportResults.TopIndex = lstEmailImportResults.Items.Count - 1;
                    }

                    if (lstEmailImportResults.Items.Count > 1000)
                    {
                        lstEmailImportResults.Items.RemoveAt(0);
                    }

                    lstEmailImportResults.Refresh();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Log error: {ex.Message}");
            }
        }

        private void LogMkgResult(string message)
        {
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action<string>(LogMkgResult), message);
                    return;
                }

                if (lstMkgResults != null)
                {
                    var timestamp = DateTime.Now.ToString("HH:mm:ss");
                    lstMkgResults.Items.Add($"[{timestamp}] {message}");

                    if (lstMkgResults.Items.Count > 0)
                    {
                        lstMkgResults.TopIndex = lstMkgResults.Items.Count - 1;
                    }

                    lstMkgResults.Refresh();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"MKG log error: {ex.Message}");
            }
        }

        private void ClearAllResults()
        {
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(ClearAllResults));
                    return;
                }

                lstEmailImportResults?.Items.Clear();
                lstMkgResults?.Items.Clear();
                lstFailedInjections?.Items.Clear();

                ShowDefaultMessages();

                Console.WriteLine("All result lists cleared");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error clearing results: {ex.Message}");
            }
        }

        private void ShowDefaultMessages()
        {
            try
            {
                if (lstEmailImportResults != null)
                {
                    lstEmailImportResults.Items.Add("=== ENHANCED EMAIL IMPORT SYSTEM ===");
                    lstEmailImportResults.Items.Add("Ready for enhanced email processing");
                    lstEmailImportResults.Items.Add("Domain intelligence loaded");
                    lstEmailImportResults.Items.Add("Real-time progress tracking enabled");
                    lstEmailImportResults.Items.Add("");
                    lstEmailImportResults.Items.Add("Click 'Email Import' to start...");
                }

                if (lstMkgResults != null)
                {
                    lstMkgResults.Items.Add("=== ENHANCED MKG PROCESSING ===");
                    lstMkgResults.Items.Add("Ready for enhanced MKG operations");
                    lstMkgResults.Items.Add("Advanced API testing capabilities");
                    lstMkgResults.Items.Add("");
                    lstMkgResults.Items.Add("Click 'MKG Test API' to run tests...");
                }

                if (lstFailedInjections != null)
                {
                    lstFailedInjections.Items.Add("=== ENHANCED ERROR TRACKING ===");
                    lstFailedInjections.Items.Add("Ready for enhanced error monitoring");
                    lstFailedInjections.Items.Add("Intelligent error analysis enabled");
                    lstFailedInjections.Items.Add("");
                    lstFailedInjections.Items.Add("Error tracking active...");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Default messages warning: {ex.Message}");
            }
        }

        private string GetDynamicProperty(dynamic obj, string propertyName)
        {
            try
            {
                var objType = obj.GetType();
                var property = objType.GetProperty(propertyName);
                if (property != null)
                {
                    var value = property.GetValue(obj);
                    return value?.ToString() ?? "";
                }
                return "";
            }
            catch
            {
                return "";
            }
        }
        #endregion

        #region Copy and Export Methods
        private void CopyListBoxToClipboard(ListBox listBox, string tabName)
        {
            try
            {
                if (listBox != null && listBox.Items.Count > 0)
                {
                    var sb = new StringBuilder();
                    sb.AppendLine($"=== {tabName} ===");
                    sb.AppendLine($"Exported at: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    sb.AppendLine();

                    foreach (var item in listBox.Items)
                    {
                        sb.AppendLine(item.ToString());
                    }

                    Clipboard.SetText(sb.ToString());

                    bool showModal = bool.Parse(ConfigurationManager.AppSettings["UI:ShowModalBoxOnCopy"] ?? "true");

                    if (showModal)
                    {
                        MessageBox.Show($"{tabName} copied to clipboard!", "Copy Tab", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    toolStripStatusLabel.Text = $"{tabName} copied to clipboard!";
                }
                else
                {
                    bool showModal = bool.Parse(ConfigurationManager.AppSettings["UI:ShowModalBoxOnCopy"] ?? "true");

                    if (showModal)
                    {
                        MessageBox.Show($"No data in {tabName} to copy. Run a test first.", "Copy Tab", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        toolStripStatusLabel.Text = $"No data in {tabName} to copy.";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error copying {tabName}: {ex.Message}", "Copy Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CopyAllTabsToClipboard()
        {
            try
            {
                if ((lstEmailImportResults?.Items.Count > 0) || (lstMkgResults?.Items.Count > 0) || (lstFailedInjections?.Items.Count > 0))
                {
                    var sb = new StringBuilder();
                    sb.AppendLine("=== MKG ELCOTEC AUTOMATION - ALL TABS EXPORT ===");
                    sb.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    sb.AppendLine($"Exported by: {Environment.UserName}");
                    sb.AppendLine();

                    if (lstEmailImportResults?.Items.Count > 0)
                    {
                        sb.AppendLine("=== EMAIL IMPORT RESULTS ===");
                        foreach (var item in lstEmailImportResults.Items)
                        {
                            sb.AppendLine(item.ToString());
                        }
                        sb.AppendLine();
                    }

                    if (lstMkgResults?.Items.Count > 0)
                    {
                        sb.AppendLine("=== MKG TEST RESULTS ===");
                        foreach (var item in lstMkgResults.Items)
                        {
                            sb.AppendLine(item.ToString());
                        }
                        sb.AppendLine();
                    }

                    if (lstFailedInjections?.Items.Count > 0)
                    {
                        sb.AppendLine("=== FAILED INJECTIONS ANALYSIS ===");
                        foreach (var item in lstFailedInjections.Items)
                        {
                            sb.AppendLine(item.ToString());
                        }
                    }

                    sb.AppendLine();
                    sb.AppendLine("=== END OF EXPORT ===");
                    sb.AppendLine($"Export completed: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

                    Clipboard.SetText(sb.ToString());

                    bool showModal = bool.Parse(ConfigurationManager.AppSettings["UI:ShowModalBoxOnCopy"] ?? "true");

                    if (showModal)
                    {
                        MessageBox.Show("All tabs copied to clipboard!", "Copy All", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    toolStripStatusLabel.Text = "All tabs copied to clipboard!";
                }
                else
                {
                    bool showModal = bool.Parse(ConfigurationManager.AppSettings["UI:ShowModalBoxOnCopy"] ?? "true");

                    if (showModal)
                    {
                        MessageBox.Show("No data in any tabs to copy. Run a test first.", "Copy All", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        toolStripStatusLabel.Text = "No data to copy.";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error copying all tabs: {ex.Message}", "Copy Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportResultsToFile()
        {
            try
            {
                var saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                saveFileDialog.DefaultExt = "txt";
                saveFileDialog.FileName = $"MKG_Results_{DateTime.Now:yyyyMMdd_HHmmss}.txt";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var sb = new StringBuilder();
                    sb.AppendLine("=== MKG ELCOTEC AUTOMATION - ALL TABS EXPORT ===");
                    sb.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    sb.AppendLine($"Exported by: {Environment.UserName}");
                    sb.AppendLine();

                    if (lstEmailImportResults?.Items.Count > 0)
                    {
                        sb.AppendLine("=== EMAIL IMPORT RESULTS ===");
                        foreach (var item in lstEmailImportResults.Items)
                        {
                            sb.AppendLine(item.ToString());
                        }
                        sb.AppendLine();
                    }

                    if (lstMkgResults?.Items.Count > 0)
                    {
                        sb.AppendLine("=== MKG TEST RESULTS ===");
                        foreach (var item in lstMkgResults.Items)
                        {
                            sb.AppendLine(item.ToString());
                        }
                        sb.AppendLine();
                    }

                    if (lstFailedInjections?.Items.Count > 0)
                    {
                        sb.AppendLine("=== FAILED INJECTIONS ANALYSIS ===");
                        foreach (var item in lstFailedInjections.Items)
                        {
                            sb.AppendLine(item.ToString());
                        }
                    }

                    sb.AppendLine();
                    sb.AppendLine("=== END OF EXPORT ===");
                    sb.AppendLine($"Export completed: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

                    File.WriteAllText(saveFileDialog.FileName, sb.ToString());

                    bool showModal = bool.Parse(ConfigurationManager.AppSettings["UI:ShowModalBoxOnCopy"] ?? "true");

                    if (showModal)
                    {
                        MessageBox.Show($"Results exported to {Path.GetFileName(saveFileDialog.FileName)}", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    toolStripStatusLabel.Text = $"Results exported to {Path.GetFileName(saveFileDialog.FileName)}";
                }
                else
                {
                    toolStripStatusLabel.Text = "Export cancelled";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting to file: {ex.Message}", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Data Conversion Methods

        /// <summary>
        /// Convert email summary to order lines - Enhanced for multiple types
        /// </summary>
        private List<OrderLine> ConvertEmailSummaryToOrderLines(EmailImportSummary importSummary)
        {
            var orderLines = new List<OrderLine>();

            try
            {
                foreach (var emailDetail in importSummary.EmailDetails)
                {
                    if (emailDetail.Orders?.Any() == true)
                    {
                        foreach (var orderData in emailDetail.Orders)
                        {
                            try
                            {
                                var orderLine = new OrderLine(
                                    orderLines.Count.ToString("D3"),
                                    GetDynamicProperty(orderData, "ArtiCode") ?? "",
                                    GetDynamicProperty(orderData, "Description") ?? "",
                                    GetDynamicProperty(orderData, "DrawingNumber") ?? "",
                                    GetDynamicProperty(orderData, "Revision") ?? "00",
                                    "",
                                    "",
                                    GetDynamicProperty(orderData, "RequestedDelivery") ?? "",
                                    GetDynamicProperty(orderData, "Description") ?? "",
                                    "",
                                    "",
                                    GetDynamicProperty(orderData, "Quantity") ?? "1",
                                    GetDynamicProperty(orderData, "Unit") ?? "PCS",
                                    GetDynamicProperty(orderData, "UnitPrice") ?? "0.00",
                                    GetDynamicProperty(orderData, "TotalPrice") ?? "0.00"
                                );

                                orderLine.PoNumber = GetDynamicProperty(orderData, "PoNumber") ?? "";
                                orderLine.DebtorNumber = emailDetail.ClientDomain ?? "";

                                orderLine.SetExtractionDetails(
                                    GetDynamicProperty(orderData, "ExtractionMethod") ?? "EMAIL_EXTRACTION",
                                    GetDynamicProperty(orderData, "ExtractionDomain") ?? emailDetail.ClientDomain
                                );

                                orderLines.Add(orderLine);
                            }
                            catch (Exception ex)
                            {
                                LogResult($"⚠️ Error converting order: {ex.Message}");
                            }
                        }
                    }
                }
                return orderLines;
            }
            catch (Exception ex)
            {
                LogResult($"❌ Error converting email summary to order lines: {ex.Message}");
                return new List<OrderLine>();
            }
        }

        /// <summary>
        /// Convert email summary to quote lines - Enhanced
        /// </summary>
        private List<QuoteLine> ConvertEmailSummaryToQuoteLines(EmailImportSummary importSummary)
        {
            var quoteLines = new List<QuoteLine>();

            try
            {
                Console.WriteLine("🔄 Converting email summary to quote lines...");

                foreach (var emailDetail in importSummary.EmailDetails)
                {
                    if (emailDetail.Orders?.Any() == true)
                    {
                        foreach (var item in emailDetail.Orders)
                        {
                            try
                            {
                                var artiCode = GetDynamicProperty(item, "ArtiCode");

                                // Check if this is a quote item (starts with QUOTE-ITEM)
                                if (!string.IsNullOrEmpty(artiCode) && artiCode.StartsWith("QUOTE-ITEM"))
                                {
                                    var quoteLine = new QuoteLine
                                    {
                                        Id = Guid.NewGuid(),
                                        ArtiCode = artiCode,
                                        Description = GetDynamicProperty(item, "Description") ?? $"Quote for {artiCode}",
                                        RfqNumber = GetDynamicProperty(item, "RfqNumber") ?? GetDynamicProperty(item, "PoNumber") ?? "RFQ-UNKNOWN",
                                        Quantity = GetDynamicProperty(item, "Quantity") ?? "1",
                                        Unit = GetDynamicProperty(item, "Unit") ?? "PCS",
                                        QuotedPrice = GetDynamicProperty(item, "QuotedPrice") ?? GetDynamicProperty(item, "UnitPrice") ?? "0.00",
                                        ExtractionMethod = GetDynamicProperty(item, "ExtractionMethod") ?? "EmailExtraction",
                                        QuoteValidUntil = DateTime.Now.AddDays(30).ToString("yyyy-MM-dd"),
                                        Priority = "Normal"
                                    };

                                    quoteLine.SetEmailDomain(emailDetail.Sender ?? "unknown@domain.com");
                                    quoteLines.Add(quoteLine);
                                    Console.WriteLine($"   ✅ Converted quote: {quoteLine.ArtiCode} (RFQ: {quoteLine.RfqNumber})");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"   ❌ Error processing quote item: {ex.Message}");
                            }
                        }
                    }
                }

                // ✅ REMOVED: Duplicate increment call - now handled in ProcessEnhancedMkgInjection
                Console.WriteLine($"✅ Converted {quoteLines.Count} quote lines from {importSummary.EmailDetails.Count} emails");
                return quoteLines;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error converting quote lines: {ex.Message}");
                return new List<QuoteLine>();
            }
        }

        /// <summary>
        /// Convert email summary to revision lines - ONLY from actual revision emails
        /// </summary>
        private List<RevisionLine> ConvertEmailSummaryToRevisionLines(EmailImportSummary importSummary)
        {
            var revisionLines = new List<RevisionLine>();

            try
            {
                Console.WriteLine("🔄 Converting email summary to revision lines...");

                foreach (var emailDetail in importSummary.EmailDetails)
                {
                    if (emailDetail.Orders?.Any() == true)
                    {
                        foreach (var item in emailDetail.Orders)
                        {
                            try
                            {
                                var currentRev = GetDynamicProperty(item, "CurrentRevision");
                                var newRev = GetDynamicProperty(item, "NewRevision");
                                var artiCode = GetDynamicProperty(item, "ArtiCode");
                                var extractionMethod = GetDynamicProperty(item, "ExtractionMethod");

                                // ✅ CRITICAL FIX: ONLY create revisions if item was extracted as a revision
                                // Must have BOTH current and new revision AND be from revision extraction
                                bool isActualRevision = !string.IsNullOrEmpty(currentRev) &&
                                                      !string.IsNullOrEmpty(newRev) &&
                                                      !string.IsNullOrEmpty(extractionMethod) &&
                                                      (extractionMethod.ToLower().Contains("revision") ||
                                                       extractionMethod.ToLower().Contains("rev"));

                                if (isActualRevision)
                                {
                                    var revisionLine = new RevisionLine
                                    {
                                        ArtiCode = artiCode ?? "",
                                        CurrentRevision = currentRev,
                                        NewRevision = newRev,
                                        RevisionReason = GetDynamicProperty(item, "RevisionReason") ??
                                                       GetDynamicProperty(item, "Reason") ?? "Revision update",
                                        ExtractionMethod = extractionMethod,
                                        DrawingNumber = GetDynamicProperty(item, "DrawingNumber") ?? artiCode ?? "",
                                        TechnicalChanges = GetDynamicProperty(item, "TechnicalChanges") ?? "Drawing revision update"
                                    };

                                    revisionLine.SetEmailDomain(emailDetail.Sender ?? "unknown@domain.com");
                                    revisionLines.Add(revisionLine);
                                    Console.WriteLine($"   ✅ Converted revision: {revisionLine.ArtiCode} ({revisionLine.CurrentRevision} → {revisionLine.NewRevision})");
                                }
                                else
                                {
                                    // ✅ DEBUG: Log why this item is NOT being converted to a revision
                                    Console.WriteLine($"   ❌ Skipping {artiCode} - NOT a revision item");
                                    Console.WriteLine($"       CurrentRev: '{currentRev}', NewRev: '{newRev}', ExtractionMethod: '{extractionMethod}'");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"   ❌ Error processing revision item: {ex.Message}");
                            }
                        }
                    }
                }

                // ✅ REMOVED: Duplicate increment call - now handled in ProcessEnhancedMkgInjection
                Console.WriteLine($"✅ Converted {revisionLines.Count} revision lines from {importSummary.EmailDetails.Count} emails");
                return revisionLines;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error converting revision lines: {ex.Message}");
                return new List<RevisionLine>();
            }
        }
        /// <summary>
        /// ✅ UPDATED: Display email import results - Include D.Emails in summary + conditional sections
        /// </summary>
        private void DisplayEmailImportResults(EmailImportSummary summary)
        {
            try
            {
                if (summary == null)
                {
                    if (lstEmailImportResults != null)
                    {
                        lstEmailImportResults.Items.Clear();
                        lstEmailImportResults.Items.Add("❌ No email import summary available");
                    }
                    return;
                }

                if (lstEmailImportResults != null)
                {
                    lstEmailImportResults.Items.Clear();

                    // Header with timestamp
                    lstEmailImportResults.Items.Add("📧 === ENHANCED EMAIL IMPORT RESULTS ===");
                    lstEmailImportResults.Items.Add($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    lstEmailImportResults.Items.Add($"User: {Environment.UserName}");
                    lstEmailImportResults.Items.Add("");

                    // ✅ UPDATED: Summary statistics - Always show all 5 lines even if 0
                    lstEmailImportResults.Items.Add("📊 === IMPORT SUMMARY ===");
                    lstEmailImportResults.Items.Add($"📧 Total Emails: {summary.TotalEmails}");
                    lstEmailImportResults.Items.Add($"✅ Successfully Processed: {summary.ProcessedEmails}");
                    lstEmailImportResults.Items.Add($"❌ Failed to Process: {summary.FailedEmails}");
                    lstEmailImportResults.Items.Add($"⏭️ Skipped Emails: {summary.SkippedEmailsCount}");
                    lstEmailImportResults.Items.Add($"🔄 Duplicate Emails: {summary.DuplicateEmailsCount}");
                    lstEmailImportResults.Items.Add("");

                    // ✅ EMAIL PROCESSING DETAILS - SHOW ONLY ONCE WITH ATTACHMENTS
                    lstEmailImportResults.Items.Add("📧 === EMAIL PROCESSING DETAILS ===");
                    lstEmailImportResults.Items.Add($"📊 Showing details for ALL {summary.EmailDetails.Count} emails:");
                    lstEmailImportResults.Items.Add("");

                    foreach (var email in summary.EmailDetails)
                    {
                        lstEmailImportResults.Items.Add($"📨 Email: {email.Subject}");
                        lstEmailImportResults.Items.Add($"   ├── From: {email.Sender}");
                        lstEmailImportResults.Items.Add($"   ├── Date: {email.ReceivedDate:yyyy-MM-dd HH:mm}");
                        lstEmailImportResults.Items.Add($"   ├── Domain: {email.ClientDomain ?? "Unknown"}");

                        // Get customer info for display
                        var customerInfo = EnhancedDomainProcessor.GetCustomerInfoForDomain(email.ClientDomain ?? "");
                        var customerName = customerInfo?.CustomerName ?? "Unknown Customer";
                        lstEmailImportResults.Items.Add($"   ├── Customer: {customerName}");
                        lstEmailImportResults.Items.Add($"   │");

                        // Show attachments
                        if (email.Attachments?.Any() == true)
                        {
                            lstEmailImportResults.Items.Add($"   ├── Attachments: {email.Attachments.Count} files");
                            foreach (var attachment in email.Attachments)
                            {
                                var sizeKB = attachment.Size / 1024.0;
                                var sizeFormatted = sizeKB < 1024 ? $"{sizeKB:F1} KB" : $"{sizeKB / 1024:F1} MB";
                                var extension = string.IsNullOrEmpty(attachment.Extension) ? "No ext" : attachment.Extension;
                                lstEmailImportResults.Items.Add($"   │   • {attachment.Name} ({extension}, {sizeFormatted})");
                            }
                        }
                        else
                        {
                            lstEmailImportResults.Items.Add($"   ├── Attachments: None");
                        }

                        lstEmailImportResults.Items.Add($"   │");

                        // Show business content extracted
                        if (email.Orders?.Any() == true)
                        {
                            var orderCount = email.Orders.Count(o => HasProperty(o, "PoNumber"));
                            var quoteCount = email.Orders.Count(o => HasProperty(o, "RfqNumber") || HasProperty(o, "QuoteNumber"));
                            var revisionCount = email.Orders.Count(o => HasProperty(o, "CurrentRevision") || HasProperty(o, "NewRevision"));

                            lstEmailImportResults.Items.Add($"   ├── Content: {orderCount} orders, {quoteCount} quotes, {revisionCount} revisions");

                            // Show individual items
                            foreach (var item in email.Orders)
                            {
                                var artiCode = GetPropertyValue(item, "ArtiCode") ?? "Unknown";
                                var itemType = "Item";

                                if (HasProperty(item, "PoNumber"))
                                    itemType = "Order";
                                else if (HasProperty(item, "RfqNumber"))
                                    itemType = "Quote";
                                else if (HasProperty(item, "CurrentRevision"))
                                    itemType = "Revision";

                                lstEmailImportResults.Items.Add($"   │   └── {itemType}: {artiCode}");
                            }
                        }
                        else
                        {
                            lstEmailImportResults.Items.Add($"   └── ⚠️ No business content extracted");
                        }
                        lstEmailImportResults.Items.Add("");
                    }

                    // ✅ CONDITIONAL: DUPLICATE EMAILS SUMMARY (only show if > 0)
                    if (summary.DuplicateEmailsCount > 0)
                    {
                        lstEmailImportResults.Items.Add("🔄 === DUPLICATE EMAILS SUMMARY ===");
                        lstEmailImportResults.Items.Add($"📊 Total Duplicates: {summary.DuplicateEmailsCount}");
                        lstEmailImportResults.Items.Add("");

                        if (summary.DuplicateEmails?.Any() == true)
                        {
                            // Group by reason for cleaner display
                            var duplicatesByReason = summary.DuplicateEmails.GroupBy(e => e.Reason ?? "Unknown reason").ToList();

                            foreach (var reasonGroup in duplicatesByReason)
                            {
                                lstEmailImportResults.Items.Add($"📋 Reason: {reasonGroup.Key} ({reasonGroup.Count()} emails)");

                                foreach (var duplicate in reasonGroup)
                                {
                                    lstEmailImportResults.Items.Add($"   🔄 {duplicate.Subject}");
                                    lstEmailImportResults.Items.Add($"      ├── From: {duplicate.Sender}");
                                    lstEmailImportResults.Items.Add($"      └── Date: {duplicate.ReceivedDate:yyyy-MM-dd HH:mm}");
                                }
                                lstEmailImportResults.Items.Add("");
                            }
                        }
                        else
                        {
                            lstEmailImportResults.Items.Add("ℹ️ No detailed duplicate information available");
                            lstEmailImportResults.Items.Add("");
                        }
                    }

                    // ✅ CONDITIONAL: SKIPPED EMAILS SUMMARY (only show if > 0)
                    if (summary.SkippedEmailsCount > 0)
                    {
                        lstEmailImportResults.Items.Add("⏭️ === SKIPPED EMAILS SUMMARY ===");
                        lstEmailImportResults.Items.Add($"📊 Total Skipped: {summary.SkippedEmailsCount}");
                        lstEmailImportResults.Items.Add("");

                        if (summary.SkippedEmails?.Any() == true)
                        {
                            // Group by reason for cleaner display
                            var skippedByReason = summary.SkippedEmails.GroupBy(e => e.Reason ?? "Unknown reason").ToList();

                            foreach (var reasonGroup in skippedByReason)
                            {
                                lstEmailImportResults.Items.Add($"📋 Reason: {reasonGroup.Key} ({reasonGroup.Count()} emails)");

                                foreach (var skipped in reasonGroup)
                                {
                                    lstEmailImportResults.Items.Add($"   ⏭️ {skipped.Subject}");
                                    lstEmailImportResults.Items.Add($"      ├── From: {skipped.Sender}");
                                    lstEmailImportResults.Items.Add($"      └── Date: {skipped.ReceivedDate:yyyy-MM-dd HH:mm}");
                                }
                                lstEmailImportResults.Items.Add("");
                            }
                        }
                        else
                        {
                            lstEmailImportResults.Items.Add("ℹ️ No detailed skip information available");
                            lstEmailImportResults.Items.Add("");
                        }
                    }

                    // ✅ CONDITIONAL: FAILED EMAILS SUMMARY (only show if > 0)
                    if (summary.FailedEmails > 0)
                    {
                        lstEmailImportResults.Items.Add("❌ === FAILED TO PROCESS SUMMARY ===");
                        lstEmailImportResults.Items.Add($"📊 Total Failed: {summary.FailedEmails}");
                        lstEmailImportResults.Items.Add("");

                        // Note: Failed emails are typically in SkippedEmails list with processing error reasons
                        var failedEmails = summary.SkippedEmails?.Where(e => e.Reason?.Contains("Processing error") == true).ToList();
                        if (failedEmails?.Any() == true)
                        {
                            foreach (var failed in failedEmails)
                            {
                                lstEmailImportResults.Items.Add($"❌ Email: {failed.Subject}");
                                lstEmailImportResults.Items.Add($"   ├── From: {failed.Sender}");
                                lstEmailImportResults.Items.Add($"   ├── Date: {failed.ReceivedDate:yyyy-MM-dd HH:mm}");
                                lstEmailImportResults.Items.Add($"   └── Reason: {failed.Reason}");
                                lstEmailImportResults.Items.Add("");
                            }
                        }
                        else
                        {
                            lstEmailImportResults.Items.Add("ℹ️ No detailed failure information available");
                            lstEmailImportResults.Items.Add("");
                        }
                    }

                    // Calculate detailed statistics for MKG injection preview
                    var totalOrderLines = 0;
                    var totalQuoteLines = 0;
                    var totalRevisionLines = 0;
                    var orderHeaders = new Dictionary<string, List<dynamic>>();
                    var quoteHeaders = new Dictionary<string, List<dynamic>>();
                    var revisionHeaders = new Dictionary<string, List<dynamic>>();

                    // Process all emails for detailed breakdown
                    foreach (var email in summary.EmailDetails)
                    {
                        if (email.Orders?.Any() == true)
                        {
                            foreach (var item in email.Orders)
                            {
                                // Classify items by type
                                if (HasProperty(item, "ArtiCode") && HasProperty(item, "PoNumber"))
                                {
                                    // This is an order
                                    var poNumber = GetPropertyValue(item, "PoNumber") ?? "UNKNOWN-PO";
                                    if (!orderHeaders.ContainsKey(poNumber))
                                        orderHeaders[poNumber] = new List<dynamic>();
                                    orderHeaders[poNumber].Add(item);
                                    totalOrderLines++;
                                }
                                else if (HasProperty(item, "RfqNumber") && HasProperty(item, "QuotedPrice"))
                                {
                                    // This is a quote
                                    var rfqNumber = GetPropertyValue(item, "RfqNumber") ?? "UNKNOWN-RFQ";
                                    if (!quoteHeaders.ContainsKey(rfqNumber))
                                        quoteHeaders[rfqNumber] = new List<dynamic>();
                                    quoteHeaders[rfqNumber].Add(item);
                                    totalQuoteLines++;
                                }
                                else if (HasProperty(item, "CurrentRevision") || HasProperty(item, "NewRevision"))
                                {
                                    // This is a revision
                                    var artiCode = GetPropertyValue(item, "ArtiCode") ?? "UNKNOWN-ARTICLE";
                                    if (!revisionHeaders.ContainsKey(artiCode))
                                        revisionHeaders[artiCode] = new List<dynamic>();
                                    revisionHeaders[artiCode].Add(item);
                                    totalRevisionLines++;
                                }
                            }
                        }
                    }

                    // Processing Performance
                    if (summary.ProcessingTime != null)
                    {
                        lstEmailImportResults.Items.Add("⚡ === PROCESSING PERFORMANCE ===");
                        if (summary.ProcessingTime is TimeSpan timeSpan)
                        {
                            lstEmailImportResults.Items.Add($"⏱️ Processing Time: {timeSpan.TotalSeconds:F1} seconds");
                            if (timeSpan.TotalSeconds > 0)
                            {
                                var emailRate = summary.ProcessedEmails / timeSpan.TotalSeconds;
                                var itemRate = (totalOrderLines + totalQuoteLines + totalRevisionLines) / timeSpan.TotalSeconds;
                                lstEmailImportResults.Items.Add($"📧 Processing Rate: {emailRate:F1} emails/second");
                                lstEmailImportResults.Items.Add($"📋 Extraction Rate: {itemRate:F1} items/second");
                            }
                        }
                        lstEmailImportResults.Items.Add("");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error displaying email import results: {ex.Message}");
                if (lstEmailImportResults != null)
                {
                    lstEmailImportResults.Items.Clear();
                    lstEmailImportResults.Items.Add($"❌ Error displaying results: {ex.Message}");
                }
            }
        }


        private bool HasProperty(dynamic obj, string propertyName)
        {
            try
            {
                var objType = obj.GetType();
                return objType.GetProperty(propertyName) != null;
            }
            catch
            {
                return false;
            }
        }

        private string GetPropertyValue(dynamic obj, string propertyName)
        {
            try
            {
                var objType = obj.GetType();
                var property = objType.GetProperty(propertyName);
                return property?.GetValue(obj)?.ToString();
            }
            catch
            {
                return null;
            }
        }
        private string ExtractMkgOrderNumber(string errorMessage)
{
    if (string.IsNullOrEmpty(errorMessage)) return "Unknown";
    
    // Look for pattern "already exists as E30250314"
    var match = System.Text.RegularExpressions.Regex.Match(errorMessage, @"exists as ([A-Z]\d+)");
    return match.Success ? match.Groups[1].Value : "Unknown";
}
        #endregion
        #region 🎯 Run History System

        // 🎯 NEW: Proper event handler for loading historical runs

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
        }

        #endregion
    }
}