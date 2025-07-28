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
using System.Text.Json;
using System.Reflection;

namespace Mkg_Elcotec_Automation.Forms
{
    
    public partial class Elcotec : Form
    {
        #region Private Fields
        private EnhancedProgressManager _enhancedProgress;
        private bool _isProcessing = false;
        private bool _isProcessingMkg = false;
        private bool _isEmailImportRunning = false;
        private bool _isMkgInjectionRunning = false;
        private DateTime _processingStartTime;
        private DateTime _lastTabUpdate = DateTime.MinValue;
        private List<MkgOrderResult> _testErrors = new List<MkgOrderResult>();
        private bool _isEmailImportTabActive = false;
        private bool _isMkgResultsTabActive = false;
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
            InitializeTabColoringCallback();
            SetupEnhancedDuplicationEventHandlers();
            Console.WriteLine("✅ Elcotec constructor completed with event wiring");
        }
        private void InitializeTabColoringCallback()
        {
            // Set up the callback for tab coloring from EmailWorkFlowService
            EmailWorkFlowService.TabColoringCallback = (tabName, isActive) =>
            {
                try
                {
                    this.Invoke((MethodInvoker)delegate
                    {
                        switch (tabName)
                        {
                            case "EmailImport":
                                _isEmailImportTabActive = isActive;
                                _isMkgResultsTabActive = false; // Turn off MKG tab when email import is active
                                if (isActive)
                                {
                                    tabControl.SelectedIndex = 0; // Switch to Email Import tab
                                }
                                break;
                            case "MkgResults":
                                _isMkgResultsTabActive = isActive;
                                _isEmailImportTabActive = false; // Turn off email import tab when MKG is active
                                if (isActive)
                                {
                                    tabControl.SelectedIndex = 1; // Switch to MKG Results tab
                                }
                                break;
                            case "Reset":
                                // Only reset workflow tabs, NOT the Failed Injections tab
                                _isEmailImportTabActive = false;
                                _isMkgResultsTabActive = false;
                                // Failed Injections tab keeps its status-based coloring
                                break;
                        }

                        // Force redraw of tabs
                        tabControl?.Invalidate();

                        Console.WriteLine($"🎨 Tab coloring updated: EmailImport={_isEmailImportTabActive}, MkgResults={_isMkgResultsTabActive}");
                    });
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error updating tab colors: {ex.Message}");
                }
            };
        }
        public void UpdateFailedInjectionsTab(MkgOrderInjectionSummary orderSummary = null,
                                           MkgQuoteInjectionSummary quoteSummary = null,
                                           MkgRevisionInjectionSummary revisionSummary = null)
        {
            try
            {
                if (lstFailedInjections == null) return;
                lstFailedInjections.Items.Clear();

                // Get statistics from progress manager
                var businessErrors = _enhancedProgress?.GetBusinessErrors() ?? 0;
                var injectionErrors = _enhancedProgress?.GetInjectionErrors() ?? 0;
                var duplicateErrors = _enhancedProgress?.GetDuplicateErrors() ?? 0;

                // Calculate total results
                var totalOrderResults = orderSummary?.OrderResults?.Count ?? 0;
                var totalQuoteResults = quoteSummary?.QuoteResults?.Count ?? 0;
                var totalRevisionResults = revisionSummary?.RevisionResults?.Count ?? 0;

                // Calculate real failures (excluding duplicates)
                var realOrderFailures = orderSummary?.OrderResults?.Count(r => !r.Success && !r.HttpStatusCode.Contains("DUPLICATE")) ?? 0;
                var realQuoteFailures = quoteSummary?.QuoteResults?.Count(r => !r.Success && !r.HttpStatusCode.Contains("DUPLICATE")) ?? 0;
                var realRevisionFailures = revisionSummary?.RevisionResults?.Count(r => !r.Success && !r.HttpStatusCode.Contains("DUPLICATE")) ?? 0;
                var totalRealFailures = realOrderFailures + realQuoteFailures + realRevisionFailures;

                // FIXED: Directly count duplicates instead of subtracting failures
                var totalMkgDuplicates = (orderSummary?.OrderResults?.Count(r => r.HttpStatusCode.Contains("DUPLICATE")) ?? 0) +
                                         (quoteSummary?.QuoteResults?.Count(r => r.HttpStatusCode.Contains("DUPLICATE")) ?? 0) +
                                         (revisionSummary?.RevisionResults?.Count(r => r.HttpStatusCode.Contains("DUPLICATE")) ?? 0);

                // Update internal flags correctly
                _hasInjectionFailures = (totalRealFailures > 0 || businessErrors > 0);
                _hasMkgDuplicatesDetected = (totalMkgDuplicates > 0);
                _totalDuplicatesInCurrentRun = totalMkgDuplicates;  // FIXED: Set to actual duplicate count for accurate yellow triggering

                // === PROCESSING SUMMARY SECTION ===
                DisplayProcessingSummary(businessErrors, injectionErrors, duplicateErrors,
                                       totalRealFailures, totalMkgDuplicates, totalOrderResults,
                                       totalQuoteResults, totalRevisionResults);

                // === DETAILED ERROR INFORMATION ===  
                // FIXED: Show detailed section if there are ANY issues (errors OR duplicates)
                if (totalRealFailures > 0 || businessErrors > 0 || injectionErrors > 0 || totalMkgDuplicates > 0)
                {
                    DisplayDetailedErrorInformation(orderSummary, quoteSummary, revisionSummary,
                                                  businessErrors, injectionErrors, totalMkgDuplicates);
                }

                // Update status labels and UI
                UpdateStatusLabelsAndFlags(totalRealFailures, businessErrors, totalMkgDuplicates);
                UpdateTabStatus();

                Console.WriteLine($"✅ Failed Injections tab updated: Real failures={totalRealFailures}, Business errors={businessErrors}, Duplicates={totalMkgDuplicates}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating failed injections tab: {ex.Message}");
            }
        }

        private void UpdateStatusLabelsAndFlags(int totalRealFailures, int businessErrors, int totalMkgDuplicates)
        {
            // FIXED: Calculate total failures including business errors
            var totalFailures = totalRealFailures + businessErrors;
            var injectionErrors = _enhancedProgress?.GetInjectionErrors() ?? 0;
            var combinedFailures = totalFailures + injectionErrors;

            // Update current run data
            if (_currentRunData != null)
            {
                _currentRunData.HasInjectionFailures = (combinedFailures > 0);
                _currentRunData.TotalFailuresAtCompletion = combinedFailures;
                _currentRunData.HasMkgDuplicatesDetected = (totalMkgDuplicates > 0);
                _currentRunData.TotalDuplicatesDetected = totalMkgDuplicates;
            }

            // FIXED: Update status label to show correct failure count in header
            if (lblFailedStatus != null)
            {
                if (combinedFailures == 0 && totalMkgDuplicates == 0)
                {
                    lblFailedStatus.Text = "✅ No issues detected - all injections successful!";
                    lblFailedStatus.ForeColor = Color.Green;
                }
                else if (combinedFailures > 0)
                {
                    // FIXED: Show total failure count including business errors
                    lblFailedStatus.Text = $"❌ {combinedFailures} failures detected" +
                                         (totalMkgDuplicates > 0 ? $" + {totalMkgDuplicates} duplicates managed" : "");
                    lblFailedStatus.ForeColor = Color.Red;
                }
                else // Only duplicates, no real failures
                {
                    lblFailedStatus.Text = $"⚠️ {totalMkgDuplicates} duplicates detected (auto-managed)";
                    lblFailedStatus.ForeColor = Color.DarkOrange;
                }
            }
        }

        private void DisplayProcessingSummary(int businessErrors, int injectionErrors, int duplicateErrors,
                                     int totalRealFailures, int totalMkgDuplicates,
                                     int totalOrderResults, int totalQuoteResults, int totalRevisionResults)
        {
            lstFailedInjections.Items.Add("══════════════════════════════════════════");
            lstFailedInjections.Items.Add("          PROCESSING SUMMARY");
            lstFailedInjections.Items.Add("══════════════════════════════════════════");
            lstFailedInjections.Items.Add($"📅 Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            lstFailedInjections.Items.Add("");

            // === OVERALL STATUS ===
            var totalInjectionFailures = injectionErrors + totalRealFailures; // Combined counter
            var overallStatus = (totalInjectionFailures > 0 || businessErrors > 0) ? "❌ FAILURES DETECTED" :
                               (totalMkgDuplicates > 0) ? "⚠️ DUPLICATES MANAGED" : "✅ ALL SUCCESSFUL";

            lstFailedInjections.Items.Add($"🎯 Overall Status: {overallStatus}");
            lstFailedInjections.Items.Add("");

            // === PROCESSING TOTALS ===
            lstFailedInjections.Items.Add("📊 PROCESSING TOTALS:");
            lstFailedInjections.Items.Add($"   • Orders Processed: {totalOrderResults}");
            lstFailedInjections.Items.Add($"   • Quotes Processed: {totalQuoteResults}");
            lstFailedInjections.Items.Add($"   • Revisions Processed: {totalRevisionResults}");
            lstFailedInjections.Items.Add($"   • Total Items: {totalOrderResults + totalQuoteResults + totalRevisionResults}");
            lstFailedInjections.Items.Add("");

            // === ERROR SUMMARY === - FIXED: Combined injection failures counter
            lstFailedInjections.Items.Add("🚨 ERROR SUMMARY:");
            lstFailedInjections.Items.Add($"   • Business Errors: {businessErrors} (Rule violations, validation failures)");
            lstFailedInjections.Items.Add($"   • Injection Failures: {totalInjectionFailures} (Technical/API failures + real failed injections)");
            lstFailedInjections.Items.Add($"   • Managed Duplicates: {totalMkgDuplicates} (Auto-skipped, not errors)");
            lstFailedInjections.Items.Add("");

            // === SUCCESS RATES === - FIXED: Use combined failure count
            var totalProcessed = totalOrderResults + totalQuoteResults + totalRevisionResults;
            if (totalProcessed > 0)
            {
                var totalFailures = businessErrors + totalInjectionFailures;
                var successfulItems = totalProcessed - totalFailures;
                var successRate = (double)successfulItems / totalProcessed * 100;

                lstFailedInjections.Items.Add("📈 SUCCESS METRICS:");
                lstFailedInjections.Items.Add($"   • Successful Injections: {successfulItems}/{totalProcessed} ({successRate:F1}%)");
                lstFailedInjections.Items.Add($"   • Failed Injections: {totalFailures}/{totalProcessed} ({(100 - successRate):F1}%)");

                if (totalMkgDuplicates > 0)
                    lstFailedInjections.Items.Add($"   • Duplicates Handled: {totalMkgDuplicates} (Automatically managed)");
            }

            lstFailedInjections.Items.Add("");
            lstFailedInjections.Items.Add("──────────────────────────────────────────");
        }


        private void DisplayDetailedErrorInformation(MkgOrderInjectionSummary orderSummary,
                                           MkgQuoteInjectionSummary quoteSummary,
                                           MkgRevisionInjectionSummary revisionSummary,
                                           int businessErrors, int injectionErrors, int totalMkgDuplicates)
        {
            lstFailedInjections.Items.Add("");
            lstFailedInjections.Items.Add("══════════════════════════════════════════");
            lstFailedInjections.Items.Add("        DETAILED ERROR INFORMATION");
            lstFailedInjections.Items.Add("══════════════════════════════════════════");
            lstFailedInjections.Items.Add("");

            // === BUSINESS ERRORS SECTION ===
            if (businessErrors > 0)
            {
                lstFailedInjections.Items.Add($"🚨 BUSINESS ERRORS ({businessErrors}):");
                lstFailedInjections.Items.Add("   These are rule violations that require attention:");
                lstFailedInjections.Items.Add("");

                DisplayBusinessErrorDetails(orderSummary, quoteSummary, revisionSummary);
                lstFailedInjections.Items.Add("");
            }

            // === INJECTION FAILURES SECTION === - NEW: Include injection errors
            if (injectionErrors > 0)
            {
                lstFailedInjections.Items.Add($"🔧 INJECTION FAILURES ({injectionErrors}):");
                lstFailedInjections.Items.Add("   These are technical/API failures that require attention:");
                lstFailedInjections.Items.Add("");

                DisplayInjectionErrorDetails(orderSummary, quoteSummary, revisionSummary);
                lstFailedInjections.Items.Add("");
            }

            // === MANAGED DUPLICATES SECTION === - NEW: Include duplicates
            if (totalMkgDuplicates > 0)
            {
                lstFailedInjections.Items.Add($"🔄 MANAGED DUPLICATES ({totalMkgDuplicates}):");
                lstFailedInjections.Items.Add("   These duplicates were automatically handled:");
                lstFailedInjections.Items.Add("");

                DisplayDuplicateDetails(orderSummary, quoteSummary, revisionSummary);
                lstFailedInjections.Items.Add("");
            }

            // === ORDER FAILURES SECTION ===
            var orderFailures = orderSummary?.OrderResults?.Where(r => !r.Success && !r.HttpStatusCode.Contains("DUPLICATE")).ToList();
            if (orderFailures?.Any() == true)
            {
                lstFailedInjections.Items.Add($"📦 ORDER FAILURES ({orderFailures.Count}):");
                foreach (var failure in orderFailures)
                {
                    lstFailedInjections.Items.Add($"   • Article: {failure.ArtiCode}");
                    lstFailedInjections.Items.Add($"     Error: {failure.ErrorMessage}");
                    lstFailedInjections.Items.Add($"     Status: {failure.HttpStatusCode}");
                    if (!string.IsNullOrEmpty(failure.PoNumber))
                        lstFailedInjections.Items.Add($"     PO Number: {failure.PoNumber}");
                    if (failure.ValidationErrors?.Any() == true)
                        lstFailedInjections.Items.Add($"     Validation Issues: {string.Join(", ", failure.ValidationErrors)}");
                    lstFailedInjections.Items.Add("");
                }
                lstFailedInjections.Items.Add("");
            }

            // === QUOTE FAILURES SECTION ===
            var quoteFailures = quoteSummary?.QuoteResults?.Where(r => !r.Success && !r.HttpStatusCode.Contains("DUPLICATE")).ToList();
            if (quoteFailures?.Any() == true)
            {
                lstFailedInjections.Items.Add($"💰 QUOTE FAILURES ({quoteFailures.Count}):");
                foreach (var failure in quoteFailures)
                {
                    lstFailedInjections.Items.Add($"   • Article: {failure.ArtiCode}");
                    lstFailedInjections.Items.Add($"     Error: {failure.ErrorMessage}");
                    lstFailedInjections.Items.Add($"     Status: {failure.HttpStatusCode}");
                    if (!string.IsNullOrEmpty(failure.RfqNumber))
                        lstFailedInjections.Items.Add($"     RFQ Number: {failure.RfqNumber}");
                    if (failure.ValidationErrors?.Any() == true)
                        lstFailedInjections.Items.Add($"     Validation Issues: {string.Join(", ", failure.ValidationErrors)}");
                    lstFailedInjections.Items.Add("");
                }
                lstFailedInjections.Items.Add("");
            }

            // === REVISION FAILURES SECTION ===
            var revisionFailures = revisionSummary?.RevisionResults?.Where(r => !r.Success && !r.HttpStatusCode.Contains("DUPLICATE")).ToList();
            if (revisionFailures?.Any() == true)
            {
                lstFailedInjections.Items.Add($"🔄 REVISION FAILURES ({revisionFailures.Count}):");
                foreach (var failure in revisionFailures)
                {
                    lstFailedInjections.Items.Add($"   • Article: {failure.ArtiCode}");
                    lstFailedInjections.Items.Add($"     Error: {failure.ErrorMessage}");
                    lstFailedInjections.Items.Add($"     Status: {failure.HttpStatusCode}");
                    if (!string.IsNullOrEmpty(failure.CurrentRevision))
                        lstFailedInjections.Items.Add($"     Current Revision: {failure.CurrentRevision}");
                    if (!string.IsNullOrEmpty(failure.NewRevision))
                        lstFailedInjections.Items.Add($"     New Revision: {failure.NewRevision}");
                    if (failure.ValidationErrors?.Any() == true)
                        lstFailedInjections.Items.Add($"     Validation Issues: {string.Join(", ", failure.ValidationErrors)}");
                    lstFailedInjections.Items.Add("");
                }
                lstFailedInjections.Items.Add("");
            }

            // === TROUBLESHOOTING SECTION ===
            lstFailedInjections.Items.Add("🔧 TROUBLESHOOTING RECOMMENDATIONS:");
            lstFailedInjections.Items.Add("   • Check MKG API connectivity and credentials");
            lstFailedInjections.Items.Add("   • Verify data format matches expected schema");
            lstFailedInjections.Items.Add("   • Review business rules and validation logic");
            lstFailedInjections.Items.Add("   • Contact system administrator for persistent issues");
            lstFailedInjections.Items.Add("");
            lstFailedInjections.Items.Add("──────────────────────────────────────────");
        }

        private void DisplayDuplicateDetails(MkgOrderInjectionSummary orderSummary,
                                    MkgQuoteInjectionSummary quoteSummary,
                                    MkgRevisionInjectionSummary revisionSummary)
        {
            var duplicatesFound = new List<string>();

            // Check orders for duplicates
            var orderDuplicates = orderSummary?.OrderResults?
                .Where(r => r.HttpStatusCode.Contains("DUPLICATE"))
                .ToList();

            if (orderDuplicates?.Any() == true)
            {
                foreach (var duplicate in orderDuplicates)
                {
                    duplicatesFound.Add($"Order {duplicate.ArtiCode}: Already exists in MKG system");
                }
            }

            // Check quotes for duplicates
            var quoteDuplicates = quoteSummary?.QuoteResults?
                .Where(r => r.HttpStatusCode.Contains("DUPLICATE"))
                .ToList();

            if (quoteDuplicates?.Any() == true)
            {
                foreach (var duplicate in quoteDuplicates)
                {
                    duplicatesFound.Add($"Quote {duplicate.ArtiCode}: Already exists in MKG system");
                }
            }

            // Check revisions for duplicates
            var revisionDuplicates = revisionSummary?.RevisionResults?
                .Where(r => r.HttpStatusCode.Contains("DUPLICATE"))
                .ToList();

            if (revisionDuplicates?.Any() == true)
            {
                foreach (var duplicate in revisionDuplicates)
                {
                    duplicatesFound.Add($"Revision {duplicate.ArtiCode}: Already exists in MKG system");
                }
            }

            // Display ALL duplicates found
            if (duplicatesFound.Any())
            {
                foreach (var duplicate in duplicatesFound)
                {
                    lstFailedInjections.Items.Add($"   • {duplicate}");
                }
                lstFailedInjections.Items.Add("   • These were automatically skipped to prevent data corruption");
                lstFailedInjections.Items.Add("   • No manual action required");
            }
            else
            {
                lstFailedInjections.Items.Add("   • No specific duplicate details available");
                lstFailedInjections.Items.Add("   • Duplicates were handled automatically during processing");
            }
        }
        private void DisplayInjectionErrorDetails(MkgOrderInjectionSummary orderSummary,
                                         MkgQuoteInjectionSummary quoteSummary,
                                         MkgRevisionInjectionSummary revisionSummary)
        {
            var injectionErrorsFound = new List<string>();

            // FIXED: Check test errors first
            var testInjectionErrors = _testErrors?.Where(r => !r.Success && r.HttpStatusCode != "BUSINESS_ERROR" && !r.HttpStatusCode.Contains("DUPLICATE")).ToList();
            if (testInjectionErrors?.Any() == true)
            {
                foreach (var error in testInjectionErrors)
                {
                    injectionErrorsFound.Add($"Order {error.ArtiCode}: {error.ErrorMessage}");
                }
            }

            // Check orders for injection errors
            var orderInjectionErrors = orderSummary?.OrderResults?
                .Where(r => !r.Success && r.HttpStatusCode != "BUSINESS_ERROR" && !r.HttpStatusCode.Contains("DUPLICATE"))
                .ToList();

            if (orderInjectionErrors?.Any() == true)
            {
                foreach (var error in orderInjectionErrors)
                {
                    injectionErrorsFound.Add($"Order {error.ArtiCode}: {error.ErrorMessage}");
                }
            }

            // Check quotes for injection errors
            var quoteInjectionErrors = quoteSummary?.QuoteResults?
                .Where(r => !r.Success && r.HttpStatusCode != "BUSINESS_ERROR" && !r.HttpStatusCode.Contains("DUPLICATE"))
                .ToList();

            if (quoteInjectionErrors?.Any() == true)
            {
                foreach (var error in quoteInjectionErrors)
                {
                    injectionErrorsFound.Add($"Quote {error.ArtiCode}: {error.ErrorMessage}");
                }
            }

            // Check revisions for injection errors
            var revisionInjectionErrors = revisionSummary?.RevisionResults?
                .Where(r => !r.Success && r.HttpStatusCode != "BUSINESS_ERROR" && !r.HttpStatusCode.Contains("DUPLICATE"))
                .ToList();

            if (revisionInjectionErrors?.Any() == true)
            {
                foreach (var error in revisionInjectionErrors)
                {
                    injectionErrorsFound.Add($"Revision {error.ArtiCode}: {error.ErrorMessage}");
                }
            }

            // Display ALL injection errors found
            if (injectionErrorsFound.Any())
            {
                foreach (var error in injectionErrorsFound)
                {
                    lstFailedInjections.Items.Add($"   • {error}");
                }
            }
            else
            {
                lstFailedInjections.Items.Add("   • No specific injection error details available");
                lstFailedInjections.Items.Add("   • Check API connectivity and technical logs");
            }
        }


        private void DisplayBusinessErrorDetails(MkgOrderInjectionSummary orderSummary,
                                       MkgQuoteInjectionSummary quoteSummary,
                                       MkgRevisionInjectionSummary revisionSummary)
        {
            var businessErrorsFound = new List<string>();

            // FIXED: Check test errors first
            var testBusinessErrors = _testErrors?.Where(r => !r.Success && r.HttpStatusCode == "BUSINESS_ERROR").ToList();
            if (testBusinessErrors?.Any() == true)
            {
                foreach (var error in testBusinessErrors)
                {
                    businessErrorsFound.Add($"Order {error.ArtiCode}: {error.ErrorMessage}");
                }
            }

            // Check orders for business errors
            var orderBusinessErrors = orderSummary?.OrderResults?
                .Where(r => !r.Success && r.HttpStatusCode == "BUSINESS_ERROR")
                .ToList();

            if (orderBusinessErrors?.Any() == true)
            {
                foreach (var error in orderBusinessErrors)
                {
                    businessErrorsFound.Add($"Order {error.ArtiCode}: {error.ErrorMessage}");
                }
            }

            // Check quotes for business errors
            var quoteBusinessErrors = quoteSummary?.QuoteResults?
                .Where(r => !r.Success && r.HttpStatusCode == "BUSINESS_ERROR")
                .ToList();

            if (quoteBusinessErrors?.Any() == true)
            {
                foreach (var error in quoteBusinessErrors)
                {
                    businessErrorsFound.Add($"Quote {error.ArtiCode}: {error.ErrorMessage}");
                }
            }

            // Check revisions for business errors
            var revisionBusinessErrors = revisionSummary?.RevisionResults?
                .Where(r => !r.Success && r.HttpStatusCode == "BUSINESS_ERROR")
                .ToList();

            if (revisionBusinessErrors?.Any() == true)
            {
                foreach (var error in revisionBusinessErrors)
                {
                    businessErrorsFound.Add($"Revision {error.ArtiCode}: {error.ErrorMessage}");
                }
            }

            // Display ALL business errors found
            if (businessErrorsFound.Any())
            {
                foreach (var error in businessErrorsFound)
                {
                    lstFailedInjections.Items.Add($"   • {error}");
                }
            }
            else
            {
                lstFailedInjections.Items.Add("   • No specific business error details available");
                lstFailedInjections.Items.Add("   • Check validation logic and business rules");
            }
        }
        private void DisplayExtractedMkgResults(MkgResultsExtractor.MkgResultsData data)
        {
            try
            {
                lstMkgResults.Items.Clear();

                // Header with extracted metadata
                lstMkgResults.Items.Add("🚀 === COMPREHENSIVE MKG INJECTION RESULTS ===");
                lstMkgResults.Items.Add($"Generated: {(data.ExportTime != DateTime.MinValue ? data.ExportTime.ToString("yyyy-MM-dd HH:mm:ss") : DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))}");
                lstMkgResults.Items.Add($"User: {data.User ?? Environment.UserName}");
                lstMkgResults.Items.Add("");

                // Enhanced Summary Section
                DisplayEnhancedSummary(data);

                // Order Results with cleaner formatting - FIXED: Handle groups correctly
                if (data.OrderGroups.Any())
                {
                    DisplayOrderResultsEnhanced(data.OrderGroups);
                }

                // Quote Results - FIXED: Handle groups correctly
                if (data.QuoteGroups.Any())
                {
                    DisplayQuoteResultsEnhanced(data.QuoteGroups);
                }

                // Revision Results - FIXED: Handle groups correctly
                if (data.RevisionGroups.Any())
                {
                    DisplayRevisionResultsEnhanced(data.RevisionGroups);
                }

                // Processing Steps
                if (data.ProcessingSteps.Any())
                {
                    lstMkgResults.Items.Add("🔄 === PROCESSING TIMELINE ===");
                    foreach (var step in data.ProcessingSteps)
                    {
                        lstMkgResults.Items.Add($"[{step.Split(']')[0].Replace("[", "")}] {string.Join("] ", step.Split(']').Skip(1))}");
                    }
                    lstMkgResults.Items.Add("");
                }

                lstMkgResults.Items.Add("=== END OF COMPREHENSIVE MKG INJECTION RESULTS ===");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error displaying enhanced MKG results: {ex.Message}");
                lstMkgResults?.Items.Add($"❌ Error displaying results: {ex.Message}");
            }
        }

        private void DisplayEnhancedSummary(MkgResultsExtractor.MkgResultsData data)
        {
            // Use the Summary property which contains the stats
            var summary = data.Summary;

            lstMkgResults.Items.Add("=== INJECTION SUMMARY ===");
            lstMkgResults.Items.Add($"Total Items Processed: {summary.TotalItemsProcessed}");
            lstMkgResults.Items.Add($"Successfully Injected: {summary.SuccessfullyInjected}");
            lstMkgResults.Items.Add($"Duplicates Skipped: {summary.DuplicatesSkipped}");
            lstMkgResults.Items.Add($"Real Failures: {summary.RealFailures}");

            if (summary.TotalItemsProcessed > 0)
            {
                var successRate = ((double)(summary.SuccessfullyInjected + summary.DuplicatesSkipped) / summary.TotalItemsProcessed * 100);
                lstMkgResults.Items.Add($"📈 Effective Success Rate: {successRate:F1}% (including duplicates as handled)");
            }
            lstMkgResults.Items.Add("");
        }
        // The DisplayOrderResultsEnhanced method should work with MkgOrderGroup (not MkgOrderItem)

        // This method works with individual MkgOrderItem objects (different from the group version above)
        private void DisplayOrderResultsEnhanced(List<MkgResultsExtractor.MkgOrderGroup> orderGroups)
        {
            lstMkgResults.Items.Add("=== ORDER INJECTION RESULTS ===");

            // Calculate totals across all groups
            var totalItems = orderGroups.SelectMany(g => g.Items).Count();
            var successCount = orderGroups.SelectMany(g => g.Items).Count(o => o.IsSuccess);
            var duplicateCount = orderGroups.SelectMany(g => g.Items).Count(o => o.IsDuplicate);
            var failureCount = orderGroups.SelectMany(g => g.Items).Count(o => !o.IsSuccess && !o.IsDuplicate);

            lstMkgResults.Items.Add($"Lines Processed: {totalItems}");
            lstMkgResults.Items.Add($"Successful: {successCount}");
            lstMkgResults.Items.Add($"Duplicates: {duplicateCount}");
            lstMkgResults.Items.Add($"Real Failures: {failureCount}");
            lstMkgResults.Items.Add($"Order Details:");

            // Show first 10 groups
            foreach (var group in orderGroups.Take(10))
            {
                lstMkgResults.Items.Add($"│    Order Group: {group.PoNumber ?? "Unknown"} ({group.Items.Count} items)");

                // Show first few items from each group
                foreach (var item in group.Items.Take(3))
                {
                    var statusIcon = item.IsSuccess ? "✅" : (item.IsDuplicate ? "🔄" : "❌");
                    var statusText = item.IsSuccess ? "SUCCESS" : (item.IsDuplicate ? "DUPLICATE_SKIPPED" : "FAILED");

                    lstMkgResults.Items.Add($"│       ├─ {item.ArticleCode} | Status:{statusText} {statusIcon} | Time:{item.ProcessedAt:HH:mm:ss}");
                    lstMkgResults.Items.Add($"│       │  ├─ 📋 {item.Description ?? "N/A"}");

                    if (!string.IsNullOrEmpty(item.MkgOrderId))
                        lstMkgResults.Items.Add($"│       │  ├─ 🆔 MKG Order: {item.MkgOrderId}");
                    if (!string.IsNullOrEmpty(item.ErrorMessage))
                        lstMkgResults.Items.Add($"│       │  └─ ⚠️ {item.ErrorMessage}");
                }

                if (group.Items.Count > 3)
                {
                    lstMkgResults.Items.Add($"│       └─ ... and {group.Items.Count - 3} more items");
                }
                lstMkgResults.Items.Add("│");
            }

            if (orderGroups.Count > 10)
            {
                lstMkgResults.Items.Add($"... and {orderGroups.Count - 10} more order groups");
            }

            lstMkgResults.Items.Add("");
        }

        // FIXED: DisplayQuoteResultsEnhanced to handle MkgQuoteGroup collections
        private void DisplayQuoteResultsEnhanced(List<MkgResultsExtractor.MkgQuoteGroup> quoteGroups)
        {
            lstMkgResults.Items.Add("=== QUOTE INJECTION RESULTS ===");

            foreach (var quoteGroup in quoteGroups.Take(10))
            {
                lstMkgResults.Items.Add($"┌── QUOTE GROUP: {quoteGroup.RfqNumber ?? "Unknown"}");

                foreach (var quote in quoteGroup.Items.Take(5))
                {
                    var statusIcon = quote.Success ? "✅" : "❌";
                    lstMkgResults.Items.Add($"│   ├── {statusIcon} {quote.ArticleCode}");
                    if (!string.IsNullOrEmpty(quote.MkgQuoteId))
                        lstMkgResults.Items.Add($"│   │   ├── MKG ID: {quote.MkgQuoteId}");
                    if (!string.IsNullOrEmpty(quote.QuotedPrice))
                        lstMkgResults.Items.Add($"│   │   ├── Price: €{quote.QuotedPrice}");
                    if (!string.IsNullOrEmpty(quote.ErrorMessage))
                        lstMkgResults.Items.Add($"│   │   └── Error: {quote.ErrorMessage}");
                }

                if (quoteGroup.Items.Count > 5)
                {
                    lstMkgResults.Items.Add($"│   └── ... and {quoteGroup.Items.Count - 5} more items");
                }
                lstMkgResults.Items.Add("│");
            }

            if (quoteGroups.Count > 10)
            {
                lstMkgResults.Items.Add($"... and {quoteGroups.Count - 10} more quote groups");
            }

            lstMkgResults.Items.Add("");
        }
        private void DisplayRevisionResultsEnhanced(List<MkgResultsExtractor.MkgRevisionGroup> revisionGroups)
        {
            lstMkgResults.Items.Add("=== REVISION INJECTION RESULTS ===");

            foreach (var revisionGroup in revisionGroups.Take(10))
            {
                lstMkgResults.Items.Add($"┌── REVISION GROUP: {revisionGroup.ArticleCode ?? "Unknown"}");

                foreach (var revision in revisionGroup.Items.Take(5))
                {
                    // Note: MkgRevisionItem might have Success property - try both Success and IsSuccess patterns
                    var success = false;
                    try
                    {
                        // Try to get Success property dynamically since definition was incomplete
                        var successProp = revision.GetType().GetProperty("Success");
                        if (successProp != null)
                        {
                            success = (bool)(successProp.GetValue(revision) ?? false);
                        }
                        else
                        {
                            // Fallback: assume success if no error message
                            var errorProp = revision.GetType().GetProperty("ErrorMessage");
                            var errorMsg = errorProp?.GetValue(revision)?.ToString();
                            success = string.IsNullOrEmpty(errorMsg);
                        }
                    }
                    catch
                    {
                        // Fallback: assume success if no error message
                        success = true;
                    }

                    var statusIcon = success ? "✅" : "❌";
                    lstMkgResults.Items.Add($"│   ├── {statusIcon} {revision.ArticleCode} | {revision.CurrentRevision} → {revision.NewRevision}");

                    // Try to get additional properties dynamically
                    try
                    {
                        var changeReasonProp = revision.GetType().GetProperty("ChangeReason");
                        var changeReason = changeReasonProp?.GetValue(revision)?.ToString();
                        if (!string.IsNullOrEmpty(changeReason))
                            lstMkgResults.Items.Add($"│   │   ├── Reason: {changeReason}");

                        var mkgRevisionIdProp = revision.GetType().GetProperty("MkgRevisionId");
                        var mkgRevisionId = mkgRevisionIdProp?.GetValue(revision)?.ToString();
                        if (!string.IsNullOrEmpty(mkgRevisionId))
                            lstMkgResults.Items.Add($"│   │   ├── MKG ID: {mkgRevisionId}");

                        var errorMessageProp = revision.GetType().GetProperty("ErrorMessage");
                        var errorMessage = errorMessageProp?.GetValue(revision)?.ToString();
                        if (!string.IsNullOrEmpty(errorMessage))
                            lstMkgResults.Items.Add($"│   │   └── Error: {errorMessage}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not access revision properties: {ex.Message}");
                    }
                }

                if (revisionGroup.Items.Count > 5)
                {
                    lstMkgResults.Items.Add($"│   └── ... and {revisionGroup.Items.Count - 5} more items");
                }
                lstMkgResults.Items.Add("│");
            }

            if (revisionGroups.Count > 10)
            {
                lstMkgResults.Items.Add($"... and {revisionGroups.Count - 10} more revision groups");
            }

            lstMkgResults.Items.Add("");
        }
        private void DisplayRevisionResultsEnhanced(List<MkgResultsExtractor.MkgRevisionItem> revisionResults)
        {
            lstMkgResults.Items.Add("=== REVISION INJECTION RESULTS ===");

            foreach (var revision in revisionResults.Take(10))
            {
                var statusIcon = revision.Success ? "✅" : "❌";
                lstMkgResults.Items.Add($"  {statusIcon} {revision.ArticleCode} | {revision.CurrentRevision} → {revision.NewRevision}");

                if (!string.IsNullOrEmpty(revision.ChangeReason))
                    lstMkgResults.Items.Add($"    Reason: {revision.ChangeReason}");

                if (!string.IsNullOrEmpty(revision.ErrorMessage))
                    lstMkgResults.Items.Add($"    Error: {revision.ErrorMessage}");
            }

            if (revisionResults.Count > 10)
                lstMkgResults.Items.Add($"... and {revisionResults.Count - 10} more revisions");

            lstMkgResults.Items.Add("");
        }

      
        private static string FormatTreeStructure(Dictionary<string, object> data)
        {
            var sb = new StringBuilder();
            var keys = data.Keys.ToList();

            for (int i = 0; i < keys.Count; i++)
            {
                var key = keys[i];
                var value = data[key];
                var isLast = i == keys.Count - 1;

                FormatTreeItem(sb, key, value, "", isLast);
            }

            return sb.ToString();
        }


        private static void FormatTreeItem(StringBuilder sb, string key, object value, string indent, bool isLast)
        {
            var prefix = isLast ? "└─" : "├─";
            var newIndent = isLast ? indent + "   " : indent + "│  ";

            if (value is Dictionary<string, object> dict)
            {
                sb.AppendLine($"{indent}{prefix} {key}:");
                var childKeys = dict.Keys.ToList();
                for (int i = 0; i < childKeys.Count; i++)
                {
                    var childKey = childKeys[i];
                    var childValue = dict[childKey];
                    var childIsLast = i == childKeys.Count - 1;
                    FormatTreeItem(sb, childKey, childValue, newIndent, childIsLast);
                }
            }
            else
            {
                sb.AppendLine($"{indent}{prefix} {key}: {value}");
            }
        }

        // Add this method to your MkgOrderController.cs or wherever the injection results are being generated:
       
        private void GenerateComprehensiveMkgResults(
      MkgOrderInjectionSummary orderSummary,
      MkgQuoteInjectionSummary quoteSummary = null,
      MkgRevisionInjectionSummary revisionSummary = null)
        {
            try
            {
                if (lstMkgResults != null)
                {
                    lstMkgResults.Items.Clear();

                    // Header
                    lstMkgResults.Items.Add("🚀 === COMPREHENSIVE MKG INJECTION RESULTS ===");
                    lstMkgResults.Items.Add($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    lstMkgResults.Items.Add($"User: {Environment.UserName}");
                    lstMkgResults.Items.Add("");

                    // Calculate totals
                    var totalItems = (orderSummary?.OrderResults?.Count ?? 0) +
                                   (quoteSummary?.QuoteResults?.Count ?? 0) +
                                   (revisionSummary?.RevisionResults?.Count ?? 0);

                    var totalSuccess = (orderSummary?.SuccessfulInjections ?? 0) +
                                     (quoteSummary?.SuccessfulInjections ?? 0) +
                                     (revisionSummary?.SuccessfulInjections ?? 0);

                    var totalDuplicates = (orderSummary?.DuplicatesDetected ?? 0) +
                                        (quoteSummary?.DuplicatesDetected ?? 0) +
                                        (revisionSummary?.DuplicatesDetected ?? 0);

                    var totalFailed = (orderSummary?.FailedInjections ?? 0) +
                                    (quoteSummary?.FailedInjections ?? 0) +
                                    (revisionSummary?.FailedInjections ?? 0) - totalDuplicates;

                    var successRate = totalItems > 0 ? (double)totalSuccess / totalItems * 100 : 0;

                    // Summary Statistics
                    lstMkgResults.Items.Add("📊 === INJECTION SUMMARY ===");
                    lstMkgResults.Items.Add($"📦 Total Items Processed: {totalItems}");
                    lstMkgResults.Items.Add($"✅ Successfully Injected: {totalSuccess}");
                    lstMkgResults.Items.Add($"🔄 Duplicates Skipped: {totalDuplicates}");
                    lstMkgResults.Items.Add($"❌ Real Failures: {totalFailed}");
                    lstMkgResults.Items.Add($"📈 Effective Success Rate: {successRate:F1}% (including duplicates as handled)");
                    lstMkgResults.Items.Add("");

                    // ORDER RESULTS
                    if (orderSummary?.OrderResults?.Any() == true)
                    {
                        DisplayOrderResultsWithRealData(orderSummary);
                    }

                    // QUOTE RESULTS
                    if (quoteSummary?.QuoteResults?.Any() == true)
                    {
                        DisplayQuoteResults(quoteSummary);
                    }

                    // REVISION RESULTS  
                    if (revisionSummary?.RevisionResults?.Any() == true)
                    {
                        DisplayRevisionResults(revisionSummary);
                    }

                    // Processing Steps
                    lstMkgResults.Items.Add("🔄 === INCREMENTAL PROCESSING STEPS ===");
                    foreach (var logEntry in GetProcessingLog())
                    {
                        lstMkgResults.Items.Add($"   {logEntry}");
                    }
                    lstMkgResults.Items.Add("");
                    lstMkgResults.Items.Add("=== END OF COMPREHENSIVE MKG INJECTION RESULTS ===");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error displaying comprehensive MKG results: {ex.Message}");
                if (lstMkgResults != null)
                {
                    lstMkgResults.Items.Add($"❌ Error displaying results: {ex.Message}");
                }
            }
        }
        private void DisplayOrderResultsWithRealData(MkgOrderInjectionSummary orderSummary)
        {
            lstMkgResults.Items.Add("📦 === ORDER INJECTION RESULTS ===");
            lstMkgResults.Items.Add($"📊 Headers Created: {orderSummary.TotalOrders}");
            lstMkgResults.Items.Add($"📋 Lines Processed: {orderSummary.OrderResults.Count}");
            lstMkgResults.Items.Add($"✅ Successful: {orderSummary.SuccessfulInjections}");
            lstMkgResults.Items.Add($"🔄 Duplicates: {orderSummary.DuplicatesDetected}");
            lstMkgResults.Items.Add($"❌ Real Failures: {orderSummary.FailedInjections - orderSummary.DuplicatesDetected}");
            lstMkgResults.Items.Add($"⏱️ Processing Time: {orderSummary.ProcessingTime.TotalSeconds:F1}s");
            lstMkgResults.Items.Add("");

            // Group by PO for detailed display
            var groupedOrders = orderSummary.OrderResults
                .GroupBy(r => r.PoNumber)
                .OrderBy(g => g.Key);

            lstMkgResults.Items.Add($"📦 Order Details ({groupedOrders.Count()} unique POs):");

            foreach (var poGroup in groupedOrders)
            {
                var successCount = poGroup.Count(r => r.Success);
                var duplicateCount = poGroup.Count(r => r.HttpStatusCode == "DUPLICATE_SKIPPED");
                var failureCount = poGroup.Count(r => !r.Success && r.HttpStatusCode != "DUPLICATE_SKIPPED");

                lstMkgResults.Items.Add($"   ├── PO: {poGroup.Key} (✅{successCount} 🔄{duplicateCount} ❌{failureCount})");

                foreach (var order in poGroup.Take(10)) // Limit to first 10 for space
                {
                    DisplaySingleOrderItem(order);
                }

                if (poGroup.Count() > 10)
                {
                    lstMkgResults.Items.Add($"   │   ... and {poGroup.Count() - 10} more items");
                }
            }
            lstMkgResults.Items.Add("");
        }

        private void DisplaySingleOrderItem(dynamic order)
        {
            var statusIcon = order.Success ? "✅" :
                           order.HttpStatusCode == "DUPLICATE_SKIPPED" ? "🔄" : "❌";

            lstMkgResults.Items.Add($"   │   {statusIcon} {order.ArtiCode} | Status:{order.HttpStatusCode} | Time:{order.ProcessedAt:HH:mm:ss}");

            // Core Data Section
            lstMkgResults.Items.Add($"   │       ├── 🎯 CORE DATA:");
            lstMkgResults.Items.Add($"   │       │   ├── Article Code: {order.ArtiCode}");
            lstMkgResults.Items.Add($"   │       │   ├── PO Number: {order.PoNumber}");
            lstMkgResults.Items.Add($"   │       │   ├── Description: {order.Description ?? "N/A"}");
            lstMkgResults.Items.Add($"   │       │   └── Success: {order.Success}");

            // MKG Identifiers Section
            lstMkgResults.Items.Add($"   │       ├── 🆔 MKG IDENTIFIERS:");
            lstMkgResults.Items.Add($"   │       │   ├── MKG Order ID: {order.MkgOrderId}");
            lstMkgResults.Items.Add($"   │       │   ├── Order Number: {order.OrderNumber ?? "N/A"}");
            lstMkgResults.Items.Add($"   │       │   ├── Line ID: {order.LineId ?? "N/A"}");
            lstMkgResults.Items.Add($"   │       │   └── Order Line ID: {order.OrderLineId ?? "N/A"}");

            // Status & Response Section
            lstMkgResults.Items.Add($"   │       ├── 📊 STATUS & RESPONSE:");
            lstMkgResults.Items.Add($"   │       │   ├── HTTP Status: {order.HttpStatusCode}");
            lstMkgResults.Items.Add($"   │       │   ├── Status Code: {order.StatusCode ?? "N/A"}");
            lstMkgResults.Items.Add($"   │       │   ├── Is Duplicate: {!order.Success && order.HttpStatusCode == "DUPLICATE_SKIPPED"}");
            lstMkgResults.Items.Add($"   │       │   └── Processed At: {order.ProcessedAt:yyyy-MM-dd HH:mm:ss}");

            // Request/Response Section - REAL DATA
            if (order.Success)
            {
                DisplayRealRequestResponse(order);
            }
            else if (order.HttpStatusCode == "DUPLICATE_SKIPPED")
            {
                lstMkgResults.Items.Add($"   │       ├── ❌ ERROR INFO:");
                lstMkgResults.Items.Add($"   │       │   └── Error: {order.ErrorMessage}");
            }

            lstMkgResults.Items.Add($"   │");
        }

        private void DisplayRealRequestResponse(dynamic order)
        {
            // Try to get real request data from various possible field names
            string actualRequest = GetOrderField(order, "RequestPayload") ??
                                  GetOrderField(order, "RequestData") ??
                                  GetOrderField(order, "Request") ??
                                  GetOrderField(order, "MkgRequest") ??
                                  "Request data not available";

            // Try to get real response data from various possible field names
            string actualResponse = GetOrderField(order, "ResponsePayload") ??
                                   GetOrderField(order, "ResponseData") ??
                                   GetOrderField(order, "Response") ??
                                   GetOrderField(order, "MkgResponse") ??
                                   GetOrderField(order, "ResponseContent") ??
                                   "Response data not available";

            // Format REQUEST as tree structure
            if (actualRequest != "Request data not available")
            {
                var requestData = ParseJsonToTreeDict(actualRequest, "");
                if (requestData != null)
                {
                    lstMkgResults.Items.Add($"   │       ├─ 📤 REQUEST:");
                    var formattedRequest = FormatTreeStructure(requestData, "   │       │  ");
                    var requestLines = formattedRequest.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var line in requestLines)
                    {
                        if (!string.IsNullOrWhiteSpace(line))
                        {
                            lstMkgResults.Items.Add($"   │       │  {line}");
                        }
                    }
                }
                else
                {
                    lstMkgResults.Items.Add($"   │       ├─ 📤 REQUEST: {TruncateData(actualRequest, 50)}");
                }
            }
            else
            {
                lstMkgResults.Items.Add($"   │       ├─ 📤 REQUEST: Not available");
            }

            // Format RESPONSE as tree structure
            if (actualResponse != "Response data not available")
            {
                var responseData = ParseJsonToTreeDict(actualResponse, "");
                if (responseData != null)
                {
                    lstMkgResults.Items.Add($"   │       └─ 📥 RESPONSE:");
                    var formattedResponse = FormatTreeStructure(responseData, "   │           ");
                    var responseLines = formattedResponse.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var line in responseLines)
                    {
                        if (!string.IsNullOrWhiteSpace(line))
                        {
                            lstMkgResults.Items.Add($"   │           {line}");
                        }
                    }
                }
                else
                {
                    lstMkgResults.Items.Add($"   │       └─ 📥 RESPONSE: {TruncateData(actualResponse, 50)}");
                }
            }
            else
            {
                lstMkgResults.Items.Add($"   │       └─ 📥 RESPONSE: Not available");
            }

            // Debug: Show available fields if data is not available (only if both are missing)
            if (actualRequest == "Request data not available" && actualResponse == "Response data not available")
            {
                var availableFields = GetAvailableFields(order);
                lstMkgResults.Items.Add($"   │       └─ 🔍 Available fields: {TruncateData(availableFields, 80)}");
            }
        }
        private Dictionary<string, object> ParseJsonToTreeDict(string jsonString, string rootName)
        {
            try
            {
                if (string.IsNullOrEmpty(jsonString)) return null;

                // Try to parse as JSON
                var jsonDoc = JsonDocument.Parse(jsonString);

                // Return the parsed structure directly without adding extra nesting
                if (jsonDoc.RootElement.ValueKind == JsonValueKind.Object)
                {
                    return JsonElementToDict(jsonDoc.RootElement) as Dictionary<string, object>;
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        private object JsonElementToDict(JsonElement element)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.Object:
                    var dict = new Dictionary<string, object>();
                    foreach (var property in element.EnumerateObject())
                    {
                        dict[property.Name] = JsonElementToDict(property.Value);
                    }
                    return dict;

                case JsonValueKind.Array:
                    var list = new Dictionary<string, object>();
                    int index = 0;
                    foreach (var item in element.EnumerateArray())
                    {
                        list[$"[{index}]"] = JsonElementToDict(item);
                        index++;
                    }
                    return list;

                default:
                    return element.ToString();
            }
        }

        private string TruncateData(string data, int maxLength)
        {
            if (string.IsNullOrEmpty(data)) return "";
            if (data.Length <= maxLength) return data;
            return data.Substring(0, maxLength) + "...";
        }

        // Update the FormatTreeStructure method to accept an indent parameter:
        private static string FormatTreeStructure(Dictionary<string, object> data, string baseIndent = "")
        {
            var sb = new StringBuilder();
            var keys = data.Keys.ToList();

            for (int i = 0; i < keys.Count; i++)
            {
                var key = keys[i];
                var value = data[key];
                var isLast = i == keys.Count - 1;

                FormatTreeItem(sb, key, value, baseIndent, isLast);
            }

            return sb.ToString();
        }
        private string GetOrderField(dynamic order, string fieldName)
        {
            try
            {
                var type = order.GetType();
                var property = type.GetProperty(fieldName);
                if (property != null)
                {
                    var value = property.GetValue(order);
                    return value?.ToString();
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        private string GetAvailableFields(dynamic order)
        {
            try
            {
                var type = order.GetType();
                var properties = type.GetProperties()
                    .Select((Func<PropertyInfo, string>)(p => p.Name))
                    .Where((Func<string, bool>)(name => name.ToLower().Contains("request") ||
                                  name.ToLower().Contains("response") ||
                                  name.ToLower().Contains("payload") ||
                                  name.ToLower().Contains("data")))
                    .ToList();
                return string.Join(", ", properties);
            }
            catch
            {
                return "Could not determine available fields";
            }
        }
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
        }// FIXES FOR ELCOTEC.CS COMPILATION ERRORS

        // ===== FIX 1: Add missing DisplayRevisionResults method =====
        private void DisplayRevisionResults(MkgRevisionInjectionSummary summary)
        {
            if (summary?.RevisionResults?.Any() != true) return;

            lstMkgResults.Items.Add("🔧 === REVISION INJECTION RESULTS ===");

            var revisionGroups = summary.RevisionResults.GroupBy(r => r.ArtiCode).ToList();

            foreach (var group in revisionGroups)
            {
                var firstRevision = group.First();
                lstMkgResults.Items.Add($"┌── REVISION: {firstRevision.ArtiCode} ({firstRevision.CurrentRevision} → {firstRevision.NewRevision})");

                if (!string.IsNullOrEmpty(firstRevision.MkgRevisionId))
                {
                    lstMkgResults.Items.Add($"│   └── MKG ID: {firstRevision.MkgRevisionId}");
                }

                foreach (var revision in group)
                {
                    var status = revision.Success ? "✅" : "❌";
                    lstMkgResults.Items.Add($"│   └── {status} {revision.FieldChanged}: {revision.OldValue} → {revision.NewValue}");

                    if (!revision.Success && !string.IsNullOrEmpty(revision.ErrorMessage))
                    {
                        lstMkgResults.Items.Add($"│       └── Error: {revision.ErrorMessage}");
                    }
                }
                lstMkgResults.Items.Add("│");
            }
            lstMkgResults.Items.Add("");
        }

        // ===== FIX 2: Add missing DisplayQuoteResults method =====
        private void DisplayQuoteResults(MkgQuoteInjectionSummary summary)
        {
            if (summary?.QuoteResults?.Any() != true) return;

            lstMkgResults.Items.Add("💰 === QUOTE INJECTION RESULTS ===");

            var quoteGroups = summary.QuoteResults.GroupBy(q => q.RfqNumber).ToList();

            foreach (var group in quoteGroups)
            {
                var firstQuote = group.First();
                lstMkgResults.Items.Add($"┌── QUOTE: RFQ {firstQuote.RfqNumber}");

                if (!string.IsNullOrEmpty(firstQuote.MkgQuoteId))
                {
                    lstMkgResults.Items.Add($"│   └── MKG ID: {firstQuote.MkgQuoteId}");
                }

                foreach (var quote in group)
                {
                    var status = quote.Success ? "✅" : "❌";
                    var price = !string.IsNullOrEmpty(quote.QuotedPrice) ? $"(€{quote.QuotedPrice})" : "";
                    lstMkgResults.Items.Add($"│   └── {status} {quote.ArtiCode} {price}");

                    if (!quote.Success && !string.IsNullOrEmpty(quote.ErrorMessage))
                    {
                        lstMkgResults.Items.Add($"│       └── Error: {quote.ErrorMessage}");
                    }
                }
                lstMkgResults.Items.Add("│");
            }
            lstMkgResults.Items.Add("");
        }

        // ===== FIX 3: Add missing GetProcessingLog method =====
        private List<string> GetProcessingLog()
        {
            try
            {
                // Try to get processing log from EmailWorkFlowService if it exists
                var emailWorkflowType = Type.GetType("Mkg_Elcotec_Automation.Services.EmailWorkFlowService");
                if (emailWorkflowType != null)
                {
                    var method = emailWorkflowType.GetMethod("GetMkgProcessingLog", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                    if (method != null)
                    {
                        var result = method.Invoke(null, null);
                        if (result is List<string> stringList)
                        {
                            return stringList;
                        }
                        else if (result is IEnumerable<string> enumerable)
                        {
                            return enumerable.ToList();
                        }
                    }
                }

                // Fallback: return empty list if service not available
                return new List<string>();
            }
            catch (Exception ex)
            {
                // Return error message if reflection fails
                return new List<string> { $"⚠️ Could not retrieve processing log: {ex.Message}" };
            }
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
                    UpdateTabStatus();
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
                    UpdateTabStatus();
                }

                Console.WriteLine("✅ UI updated with detailed historical data");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error displaying detailed historical data: {ex.Message}");
            }
        }

     
        private void DisplayDetailedFailureBreakdown(MkgOrderInjectionSummary orderSummary,
                                           MkgQuoteInjectionSummary quoteSummary,
                                           MkgRevisionInjectionSummary revisionSummary)
        {
            // Order failures
            if (orderSummary?.OrderResults != null)
            {
                var realOrderFailures = orderSummary.OrderResults.Where(r =>
                    !r.Success &&
                    !r.HttpStatusCode.Contains("DUPLICATE") &&
                    !r.ErrorMessage?.Contains("DUPLICATE") == true).ToList();

                if (realOrderFailures.Any())
                {
                    lstFailedInjections.Items.Add($"📦 ORDER SYSTEM FAILURES ({realOrderFailures.Count}):");
                    foreach (var failure in realOrderFailures.Take(10))
                    {
                        lstFailedInjections.Items.Add($"   • {failure.ArtiCode}: {failure.ErrorMessage}");
                        if (!string.IsNullOrEmpty(failure.HttpStatusCode))
                            lstFailedInjections.Items.Add($"     HTTP: {failure.HttpStatusCode}");
                    }
                    if (realOrderFailures.Count > 10)
                        lstFailedInjections.Items.Add($"   ... and {realOrderFailures.Count - 10} more order failures");
                    lstFailedInjections.Items.Add("");
                }
            }

            // Quote failures
            if (quoteSummary?.QuoteResults != null)
            {
                var realQuoteFailures = quoteSummary.QuoteResults.Where(r =>
                    !r.Success &&
                    !r.HttpStatusCode.Contains("DUPLICATE") &&
                    !r.ErrorMessage?.Contains("DUPLICATE") == true).ToList();

                if (realQuoteFailures.Any())
                {
                    lstFailedInjections.Items.Add($"💰 QUOTE SYSTEM FAILURES ({realQuoteFailures.Count}):");
                    foreach (var failure in realQuoteFailures.Take(10))
                    {
                        lstFailedInjections.Items.Add($"   • {failure.ArtiCode}: {failure.ErrorMessage}");
                        if (!string.IsNullOrEmpty(failure.HttpStatusCode))
                            lstFailedInjections.Items.Add($"     HTTP: {failure.HttpStatusCode}");
                    }
                    if (realQuoteFailures.Count > 10)
                        lstFailedInjections.Items.Add($"   ... and {realQuoteFailures.Count - 10} more quote failures");
                    lstFailedInjections.Items.Add("");
                }
            }

            // Revision failures
            if (revisionSummary?.RevisionResults != null)
            {
                var realRevisionFailures = revisionSummary.RevisionResults.Where(r =>
                    !r.Success &&
                    !r.HttpStatusCode.Contains("DUPLICATE") &&
                    !r.ErrorMessage?.Contains("DUPLICATE") == true).ToList();

                if (realRevisionFailures.Any())
                {
                    lstFailedInjections.Items.Add($"🔧 REVISION SYSTEM FAILURES ({realRevisionFailures.Count}):");
                    foreach (var failure in realRevisionFailures.Take(10))
                    {
                        lstFailedInjections.Items.Add($"   • {failure.ArtiCode}: {failure.ErrorMessage}");
                        if (!string.IsNullOrEmpty(failure.HttpStatusCode))
                            lstFailedInjections.Items.Add($"     HTTP: {failure.HttpStatusCode}");
                    }
                    if (realRevisionFailures.Count > 10)
                        lstFailedInjections.Items.Add($"   ... and {realRevisionFailures.Count - 10} more revision failures");
                    lstFailedInjections.Items.Add("");
                }
            }
        }


        // Modify UpdateTabStatus to check this flag:
        public void UpdateTabStatus()
        {
            try
            {
                if (_isProcessingMkg && DateTime.Now - _lastTabUpdate < TimeSpan.FromSeconds(2))
                {
                    return;
                }

                _lastTabUpdate = DateTime.Now;

                var currentStatus = GetCurrentTabStatus();

                if (tabControl != null)
                {
                    foreach (TabPage tab in tabControl.TabPages)
                    {
                        if (tab.Text.Contains("Failed") || tab.Text.Contains("Error"))
                        {
                            // 🔧 KEY FIX: DON'T set tab.BackColor - this affects content area
                            // Instead, rely on custom drawing in TabControl_DrawItem

                            // Keep the content area white/normal
                            tab.BackColor = SystemColors.Control; // Normal background for content
                            tab.ForeColor = SystemColors.ControlText; // Normal text for content

                            // The tab header color will be handled by TabControl_DrawItem
                            break;
                        }
                    }

                    // Force redraw of tab headers only
                    tabControl.Invalidate();
                }

                var statusText = _isEmailImportRunning ? "Processing emails..." :
                                currentStatus == TabStatus.Red ? "Injection failures detected" :
                                currentStatus == TabStatus.Yellow ? "Duplicates detected" :
                                currentStatus == TabStatus.Green ? "All systems working" : "Ready";

                Console.WriteLine($"🎨 Tab status updated: {statusText}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating tab status: {ex.Message}");
            }
        }
        private void TabControl_DrawItem(object sender, DrawItemEventArgs e)
        {
            try
            {
                TabControl tc2 = (TabControl)sender;
                TabPage tp2 = tc2.TabPages[e.Index];
                Rectangle tabRect2 = tc2.GetTabRect(e.Index);

                // Default colors for tab header
                Color bgColor = SystemColors.Control;
                Color textColor = SystemColors.ControlText;

                // Debug output
                Console.WriteLine($"🎨 Drawing tab: {tp2.Text}, EmailImportActive={_isEmailImportTabActive}, MkgResultsActive={_isMkgResultsTabActive}");

                // NEW: Apply workflow-based coloring for Email Import and MKG Results tabs
                if (tp2.Text.Contains("Email Import") && _isEmailImportTabActive)
                {
                    Console.WriteLine("🎨 Setting Email Import tab to BLUE");
                    bgColor = Color.Blue;
                    textColor = Color.White;
                }
                else if (tp2.Text.Contains("MKG Results") && _isMkgResultsTabActive)
                {
                    Console.WriteLine("🎨 Setting MKG Results tab to BLUE");
                    bgColor = Color.Blue;
                    textColor = Color.White;
                }
                // EXISTING: Apply custom colors to Failed Injections tab HEADER based on status
                else if (tp2.Text.Contains("Failed"))
                {
                    var status = GetCurrentTabStatus();
                    switch (status)
                    {
                        case TabStatus.Red:
                            bgColor = Color.Red;
                            textColor = Color.White;
                            break;
                        case TabStatus.Yellow:
                            bgColor = Color.Orange;
                            textColor = Color.White;
                            break;
                        case TabStatus.Green:
                            bgColor = Color.Green;
                            textColor = Color.White;
                            break;
                    }
                }

                // Draw tab header background
                using (var brush = new SolidBrush(bgColor))
                    e.Graphics.FillRectangle(brush, tabRect2);

                // Draw tab header text
                TextRenderer.DrawText(e.Graphics, tp2.Text, tc2.Font, tabRect2, textColor,
                                     TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in TabControl_DrawItem: {ex.Message}");
                e.DrawBackground();
            }
        }
        // Modify the DuplicationService event handlers to check the flag:
        // 🔧 FIXED: Proper status priority logic
        private TabStatus GetCurrentTabStatus()
        {
            // Get injection errors from progress manager
            var injectionErrors = _enhancedProgress?.GetInjectionErrors() ?? 0;
            var businessErrors = _enhancedProgress?.GetBusinessErrors() ?? 0;

            // DEBUG: Check the duplicate flags
            Console.WriteLine($"🔍 DEBUG Tab Status: _hasMkgDuplicatesDetected={_hasMkgDuplicatesDetected}, _totalDuplicatesInCurrentRun={_totalDuplicatesInCurrentRun}, _isProcessingMkg={_isProcessingMkg}");

            // RED takes highest priority - for injection failures OR business errors
            if (_hasInjectionFailures || injectionErrors > 0 || businessErrors > 0)
            {
                Console.WriteLine($"🔴 Tab status: RED - failures detected (injection: {injectionErrors}, business: {businessErrors}, hasInjectionFailures: {_hasInjectionFailures})");
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
                // 🚨 CRITICAL BUSINESS ERRORS - Financial Protection - RED tab
                DuplicationService.OnCriticalBusinessErrors += (count) =>
                {
                    try
                    {
                        Console.WriteLine($"🚨 CRITICAL BUSINESS ERRORS: {count} detected - immediate action required");

                        // Only set injection failures flag during MKG processing, not email import
                        if (_isProcessingMkg)
                        {
                            _hasInjectionFailures = true; // ✅ CORRECT: Force RED tab for business errors
                            UpdateTabStatus();
                        }

                        // Update current run data
                        if (_currentRunData != null)
                        {
                            _currentRunData.HasInjectionFailures = true;
                            _currentRunData.TotalFailuresAtCompletion += count;
                        }

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

                        // Only set injection failures flag during MKG processing, not email import
                        if (_isProcessingMkg)
                        {
                            _hasInjectionFailures = true; // ✅ CORRECT: Force RED tab for financial risks
                            UpdateTabStatus();
                        }

                        // Update current run data
                        if (_currentRunData != null)
                        {
                            _currentRunData.HasInjectionFailures = true;
                            _currentRunData.TotalFailuresAtCompletion++;
                        }

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

                        // Only update tab status during MKG processing, not email import
                        if (_isProcessingMkg)
                        {
                            _hasMkgDuplicatesDetected = true; // ✅ CORRECT: Only set duplicate flag
                            _totalDuplicatesInCurrentRun += count;
                            UpdateTabStatus();
                        }

                        // Update current run data - duplicates are NOT failures
                        if (_currentRunData != null)
                        {
                            _currentRunData.HasMkgDuplicatesDetected = true;
                            _currentRunData.TotalDuplicatesDetected += count;
                            _currentRunData.LastDuplicateDetectionUpdate = DateTime.Now;
                        }

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

                Console.WriteLine($"🎨 Applied historical tab colors: Errors={hasErrors}, Duplicates={hasDuplicates}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating tab colors: {ex.Message}");
            }
        }
        
        // 🎯 Helper method to apply tab colors (if not already present)
       
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

                // Reset tab colors when workflow is complete
                EmailWorkFlowService.ResetAllTabColors();

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
                EmailWorkFlowService.ResetAllTabColors();
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
                _isProcessingMkg = true;
                Console.WriteLine("🚀 Starting ENHANCED MKG injection with real-time updates...");

                // 🎯 NEW: Switch from Email Import tab to MKG Results tab
                EmailWorkFlowService.SetMkgResultsTabActive();

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
                            _enhancedProgress
                        );
                        totalInjected += orderSummary.SuccessfulInjections;
                        totalFailed += orderSummary.FailedInjections;
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
                    }
                }
                ForceBusinessError();
                ForceInjectionError();
                
                Console.WriteLine("🔍 COUNTING INJECTION ERRORS:");

                if (orderSummary?.OrderResults != null)
                {
                    var realFailures = orderSummary.OrderResults.Where(r => !r.Success &&
                        !(r.HttpStatusCode?.Contains("DUPLICATE") == true ||
                          r.ErrorMessage?.Contains("DUPLICATE") == true ||
                          r.ErrorMessage?.Contains("already exists") == true)).ToList();

                    Console.WriteLine($"📊 Real order failures found: {realFailures.Count}");

                    foreach (var realFailure in realFailures)
                    {
                        _enhancedProgress?.IncrementInjectionErrors();
                        Console.WriteLine($"🔧 Incremented I:Errors counter for {realFailure.ArtiCode}");
                    }
                }

                if (quoteSummary?.QuoteResults != null)
                {
                    var realQuoteFailures = quoteSummary.QuoteResults.Where(r => !r.Success &&
                        !(r.HttpStatusCode?.Contains("DUPLICATE") == true ||
                          r.ErrorMessage?.Contains("DUPLICATE") == true ||
                          r.ErrorMessage?.Contains("already exists") == true)).ToList();

                    foreach (var realFailure in realQuoteFailures)
                    {
                        _enhancedProgress?.IncrementInjectionErrors();
                        Console.WriteLine($"🔧 Incremented I:Errors counter for quote {realFailure.ArtiCode}");
                    }
                }

                if (revisionSummary?.RevisionResults != null)
                {
                    var realRevisionFailures = revisionSummary.RevisionResults.Where(r => !r.Success &&
                        !(r.HttpStatusCode?.Contains("DUPLICATE") == true ||
                          r.ErrorMessage?.Contains("DUPLICATE") == true ||
                          r.ErrorMessage?.Contains("already exists") == true)).ToList();

                    foreach (var realFailure in realRevisionFailures)
                    {
                        _enhancedProgress?.IncrementInjectionErrors();
                        Console.WriteLine($"🔧 Incremented I:Errors counter for revision {realFailure.ArtiCode}");
                    }
                }

                // Generate comprehensive results and display using extractor
                GenerateComprehensiveMkgResults(orderSummary, quoteSummary, revisionSummary);

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
            finally
            {
                _isProcessingMkg = false;
                UpdateTabStatus();
            }
        }
        private void LogMkgResultRealTime(string message)
        {
            try
            {
                // 🎯 SHOW REAL-TIME in lstMkgResults (keep this feature)
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

                // 🎯 ALSO log to MKG processing log for collection at end
                if (!string.IsNullOrEmpty(message))
                {
                    EmailWorkFlowService.LogMkgWorkflow(message);
                }

                // Also log to console for debugging
                Console.WriteLine(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in real-time logging: {ex.Message}");
            }
        }



        #endregion

        #region Display Methods
       
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
                    lstMkgResults.Items.Add("");

                    // 🎯 ADD INCREMENTAL STEPS HERE AT THE END OF QUOTES
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
                    var items = new List<string>();

                    foreach (var item in listBox.Items)
                    {
                        // Clean the text - remove emojis and special formatting that break JSON
                        string cleanText = item.ToString()
                            .Replace("📊", "")
                            .Replace("✅", "")
                            .Replace("❌", "")
                            .Replace("🔄", "")
                            .Replace("📧", "")
                            .Replace("💾", "")
                            .Replace("🎯", "")
                            .Replace("📦", "")
                            .Replace("📋", "")
                            .Replace("💰", "")
                            .Replace("⚠️", "")
                            .Replace("⏭️", "")
                            .Replace("🔧", "")
                            .Replace("💱", "")
                            .Trim();

                        if (!string.IsNullOrWhiteSpace(cleanText))
                        {
                            items.Add(cleanText);
                        }
                    }

                    var tabData = new
                    {
                        timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        tabName = tabName,
                        itemCount = items.Count,
                        items = items.ToArray()
                    };

                    var options = new JsonSerializerOptions
                    {
                        WriteIndented = true,
                        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                    };

                    string jsonContent = JsonSerializer.Serialize(tabData, options);
                    Clipboard.SetText(jsonContent);

                    bool showModal = bool.Parse(ConfigurationManager.AppSettings["UI:ShowModalBoxOnCopy"] ?? "true");

                    if (showModal)
                    {
                        MessageBox.Show($"{tabName} copied as valid JSON to clipboard!", "Copy Tab", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    toolStripStatusLabel.Text = $"{tabName} copied as JSON to clipboard!";
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

        private string[] GetCleanItemsFromListBox(ListBox listBox)
        {
            var items = new List<string>();

            foreach (var item in listBox.Items)
            {
                // Clean the text - remove emojis and special formatting that break JSON
                string cleanText = item.ToString()
                    .Replace("📊", "")
                    .Replace("✅", "")
                    .Replace("❌", "")
                    .Replace("🔄", "")
                    .Replace("📧", "")
                    .Replace("💾", "")
                    .Replace("🎯", "")
                    .Replace("📦", "")
                    .Replace("📋", "")
                    .Replace("💰", "")
                    .Replace("⚠️", "")
                    .Replace("⏭️", "")
                    .Replace("🔧", "")
                    .Replace("💱", "")
                    .Replace("⏭️", "")
                    .Trim();

                if (!string.IsNullOrWhiteSpace(cleanText))
                {
                    items.Add(cleanText);
                }
            }

            return items.ToArray();
        }

        // Replace the existing CopyAllTabsToClipboard method with this JSON version:

        private void CopyAllTabsToClipboard()
        {
            try
            {
                var allTabsData = new
                {
                    timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    tabs = new
                    {
                        emailImport = GetTabContentAsJson(lstEmailImportResults, "Email Import"),
                        mkgResults = GetTabContentAsJson(lstMkgResults, "MKG Results"),
                        failedInjections = GetTabContentAsJson(lstFailedInjections, "Failed Injections")
                    }
                };

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };

                string jsonContent = JsonSerializer.Serialize(allTabsData, options);

                if (!string.IsNullOrEmpty(jsonContent))
                {
                    Clipboard.SetText(jsonContent);

                    bool showModal = bool.Parse(ConfigurationManager.AppSettings["UI:ShowModalBoxOnCopy"] ?? "true");

                    if (showModal)
                    {
                        MessageBox.Show("All tabs copied as valid JSON to clipboard!", "Copy All", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    toolStripStatusLabel.Text = "All tabs copied as JSON to clipboard!";
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

        private object GetTabContentAsJson(ListBox listBox, string tabName)
        {
            if (listBox == null || listBox.Items.Count == 0)
            {
                return new { tabName = tabName, items = new string[0], status = "empty" };
            }

            var items = new List<string>();
            foreach (var item in listBox.Items)
            {
                // Clean the text - remove emojis and special formatting
                string cleanText = item.ToString()
                    .Replace("📊", "")
                    .Replace("✅", "")
                    .Replace("❌", "")
                    .Replace("🔄", "")
                    .Replace("📧", "")
                    .Replace("💾", "")
                    .Replace("🎯", "")
                    .Replace("📦", "")
                    .Replace("📋", "")
                    .Replace("💰", "")
                    .Replace("⚠️", "")
                    .Replace("⏭️", "")
                    .Trim();

                if (!string.IsNullOrWhiteSpace(cleanText))
                {
                    items.Add(cleanText);
                }
            }

            return new
            {
                tabName = tabName,
                items = items.ToArray(),
                itemCount = items.Count,
                status = "populated"
            };
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

                    // 🎯 NEW: ADD INCREMENTAL PROCESSING STEPS
                    var processingSteps = EmailWorkFlowService.GetProcessingLog();
                    if (processingSteps?.Any() == true)
                    {
                        lstEmailImportResults.Items.Add("🔄 === INCREMENTAL PROCESSING STEPS ===");
                        lstEmailImportResults.Items.Add($"📊 Total Steps: {processingSteps.Count}");
                        lstEmailImportResults.Items.Add("");

                        foreach (var step in processingSteps)
                        {
                            lstEmailImportResults.Items.Add($"   {step}");
                        }
                        lstEmailImportResults.Items.Add("");
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
        public void ForceInjectionError()
        {
            _enhancedProgress?.IncrementInjectionErrors();
            _hasInjectionFailures = true;

            if (_currentRunData != null)
            {
                _currentRunData.HasInjectionFailures = true;
                _currentRunData.TotalFailuresAtCompletion++;
            }

            // Create a test order result with injection error
            var testInjectionError = new MkgOrderResult
            {
                ArtiCode = "TEST-INJECTION-002",
                PoNumber = "PO-TEST-INJ",
                Success = false,
                ErrorMessage = "🧪 TEST: API connection timeout - Server unreachable",
                HttpStatusCode = "500",
                ValidationErrors = new List<string> { "Network timeout", "Retry limit exceeded" },
                ProcessedAt = DateTime.Now
            };

            // Add to global test errors list
            _testErrors.Add(testInjectionError);

            UpdateTabStatus();

            // Immediately update the display
            UpdateFailedInjectionsTab(_lastOrderSummary, _lastQuoteSummary, _lastRevisionSummary);

            Console.WriteLine("🔧 FORCED 1 INJECTION ERROR WITH TEST DATA");
        }

        public void ForceBusinessError()
        {
            _enhancedProgress?.IncrementBusinessErrors("TEST-Business", "Forced business rule violation for testing");
            _hasInjectionFailures = true;

            if (_currentRunData != null)
            {
                _currentRunData.HasInjectionFailures = true;
                _currentRunData.TotalFailuresAtCompletion++;
            }

            // Create a test order result with business error
            var testBusinessError = new MkgOrderResult
            {
                ArtiCode = "TEST-BUSINESS-001",
                PoNumber = "PO-TEST-BIZ",
                Success = false,
                ErrorMessage = "🧪 TEST: Business rule violation - Price exceeds customer limit",
                HttpStatusCode = "BUSINESS_ERROR",
                ValidationErrors = new List<string> { "Price validation failed", "Customer credit limit exceeded" },
                ProcessedAt = DateTime.Now
            };

            // Add to global test errors list
            _testErrors.Add(testBusinessError);

            UpdateTabStatus();

            // Immediately update the display
            UpdateFailedInjectionsTab(_lastOrderSummary, _lastQuoteSummary, _lastRevisionSummary);

            Console.WriteLine("🚨 FORCED 1 BUSINESS ERROR WITH TEST DATA");
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