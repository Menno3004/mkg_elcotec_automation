using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph.Models;
using Mkg_Elcotec_Automation.Services;
using Mkg_Elcotec_Automation.Models;
using Mkg_Elcotec_Automation.Models.EmailModels;
using Mkg_Elcotec_Automation.Controllers;
using System.Text.RegularExpressions;
using System.Text.Json;

namespace Mkg_Elcotec_Automation.Services
{

    /// <summary>
    /// Complete Email Workflow Service with MKG Integration and Live Statistics
    /// 🎯 FIXED: Now includes complete workflow with email import + MKG injection
    /// </summary>
    public static class EmailWorkFlowService
    {
        private static List<string> _processingLog = new List<string>();
        public static void ClearProcessingLog() => _processingLog.Clear();
        private static List<string> _mkgProcessingLog = new List<string>();
        public static void ClearMkgProcessingLog() => _mkgProcessingLog.Clear();
        public static List<string> GetMkgProcessingLog() => new List<string>(_mkgProcessingLog);
        public static List<string> GetProcessingLog() => new List<string>(_processingLog);
        public static Action<string, bool> TabColoringCallback { get; set; }
        private static void LogWorkflow(string message)
        {
            var logEntry = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
            _processingLog.Add(logEntry);
            Console.WriteLine($"[EMAIL_WORKFLOW] {logEntry}");
        }
        public static void LogMkgWorkflow(string message)
        {
            var logEntry = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
            _mkgProcessingLog.Add(logEntry);
            Console.WriteLine($"[MKG_WORKFLOW] {logEntry}");
        }

        public static void SetEmailImportTabActive()
        {
            TabColoringCallback?.Invoke("EmailImport", true);
        }
        public static void SetMkgResultsTabActive()
        {
            TabColoringCallback?.Invoke("MkgResults", true);
        }
        public static void ResetAllTabColors()
        {
            TabColoringCallback?.Invoke("Reset", false);
        }
        public static async Task<EmailImportSummary> ImportEmailsAsync(
    EnhancedProgressManager progressManager = null,
    Action<string> logCallback = null)
        {
            var summary = new EmailImportSummary();
            ClearProcessingLog();
            var trackingData = new TrackingData();

            try
            {
                LogWorkflow("=== EMAIL IMPORT STARTED ===");
                logCallback?.Invoke("🚀 Starting email import process...");
                SetEmailImportTabActive();
                // 🎯 FIXED: Use StartNewSession for initial email import
                progressManager?.StartNewSession("Email Import", 100);

                progressManager?.UpdateDebugConsole("🚀 Starting email import...");
                logCallback?.Invoke("🔗 Connecting to Microsoft Graph API...");
                await Task.Delay(1);

                // Step 1: Initialize GraphHandler with credentials from config
                progressManager?.UpdateProgress(10, 100, "Connecting to Microsoft Graph...");
                progressManager?.UpdateDebugConsole("🔗 Connecting to Microsoft Graph...");

                var tenantId = ConfigurationManager.AppSettings["Email:TenantId"];
                var clientId = ConfigurationManager.AppSettings["Email:ClientId"];
                var clientSecret = ConfigurationManager.AppSettings["Email:ClientSecret"];
                var userEmail = ConfigurationManager.AppSettings["Email:User"];

                if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId) ||
                    string.IsNullOrEmpty(clientSecret) || string.IsNullOrEmpty(userEmail))
                {
                    throw new InvalidOperationException("Email configuration missing in App.config. " +
                        "Please check Email:TenantId, Email:ClientId, Email:ClientSecret, and Email:User settings.");
                }

                LogWorkflow($"📧 Using user email: {userEmail}");
                logCallback?.Invoke($"📧 Authenticating as: {userEmail}");

                // Step 2: Create GraphHandler with credentials (authentication happens in constructor)
                var graphHandler = new GraphHandler(tenantId, clientId, clientSecret);
                LogWorkflow("✅ GraphHandler initialized successfully");
                logCallback?.Invoke("✅ Microsoft Graph connection established");

                try
                {
                    // Step 3: Get emails from Microsoft Graph using existing method
                    progressManager?.UpdateProgress(20, 100, "Retrieving emails from Microsoft Graph...");
                    progressManager?.UpdateDebugConsole("📧 Retrieving emails from Microsoft Graph...");
                    logCallback?.Invoke("📧 Fetching emails from inbox...");

                    var emailList = await graphHandler.GetAllEmailsFromUser(userEmail);

                    if (!emailList.Any())
                    {
                        LogWorkflow("⚠️ No emails found in the user's inbox");
                        progressManager?.CompleteOperation("No emails found to process");
                        logCallback?.Invoke("⚠️ No emails found in inbox");
                        return summary;
                    }

                    // 🎯 FIX: Don't set total all at once, let it increment as we process
                    summary.TotalEmails = emailList.Count();
                    LogWorkflow($"📧 Found {emailList.Count()} emails to process");
                    logCallback?.Invoke($"📧 Found {emailList.Count()} emails to process");

                    // Step 4: Process each email with enhanced duplicate and domain filtering
                    progressManager?.UpdateProgress(30, 100, "Processing emails for business content...");
                    progressManager?.UpdateDebugConsole($"🎯 Processing emails: 0/{emailList.Count()}");
                    logCallback?.Invoke($"🎯 Starting to process {emailList.Count()} emails...");

                    int processedCount = 0;

                    // FIXED EMAIL PROCESSING LOOP - This is the main issue with S.Emails counter

                    foreach (var email in emailList)
                    {
                        try
                        {
                            processedCount++;
                            progressManager?.IncrementTotalEmails(); // T.Emails++

                            var emailProgress = 30 + (processedCount * 40 / emailList.Count());
                            progressManager?.UpdateProgress(emailProgress, 100, $"Processing email {processedCount}/{emailList.Count()}...");

                            var emailSubjectPreview = (email.Subject?.Length > 40) ? email.Subject.Substring(0, 40) + "..." : email.Subject ?? "No Subject";
                            progressManager?.UpdateDebugConsole($"🎯 Processing email {processedCount}/{emailList.Count()}: {emailSubjectPreview}");
                            logCallback?.Invoke($"📧 [{processedCount}/{emailList.Count()}] Processing: {email.Subject}");

                            // STEP 1: Check for duplicates FIRST
                            var emailKey = GenerateEmailKey(email);
                            if (trackingData.ProcessedEmailIds.Contains(emailKey))
                            {
                                LogWorkflow($"🔄 Duplicate email detected: {email.Subject}");
                                logCallback?.Invoke($"🔄 Duplicate email skipped: {emailSubjectPreview}");
                                logCallback?.Invoke(""); // RESTORE LINE SPACING

                                // Duplicates go to D.Emails
                                progressManager?.IncrementDuplicateEmails();
                                trackingData.DuplicateEmailsCount++;

                                summary.DuplicateEmails.Add(new DuplicateEmailDetail
                                {
                                    Subject = email.Subject ?? "No Subject",
                                    Sender = email.From?.EmailAddress?.Address ?? "Unknown Sender",
                                    Reason = "Duplicate content from same client",
                                    ReceivedDate = email.ReceivedDateTime?.DateTime ?? DateTime.Now
                                });
                                continue; // Skip this email completely
                            }


                            var senderDomain = ExtractDomain(email.From?.EmailAddress?.Address ?? "");
                            var customerInfo = EnhancedDomainProcessor.GetCustomerInfoForDomain(senderDomain);

                            if (customerInfo == null)
                            {
                                LogWorkflow($"⏭️ Email skipped: Unsupported domain {senderDomain}");
                                logCallback?.Invoke($"⏭️ Skipped - unsupported domain: {senderDomain}");
                                logCallback?.Invoke(""); // RESTORE LINE SPACING
                                progressManager?.IncrementSkippedEmails();
                                trackingData.SkippedEmailsCount++;
                                summary.SkippedEmails.Add(new SkippedEmailDetail
                                {
                                    Subject = email.Subject ?? "No Subject",
                                    Sender = email.From?.EmailAddress?.Address ?? "Unknown Sender",
                                    Reason = $"Unsupported domain: {senderDomain}",
                                    ReceivedDate = email.ReceivedDateTime?.DateTime ?? DateTime.Now
                                });
                                continue; // Skip this email completely
                            }

                            // STEP 2.5: Lite v1.0 - Only process Weir customers
                            if (!EnhancedDomainProcessor.IsWeirCustomer(customerInfo))
                            {
                                LogWorkflow($"⏭️ Email skipped: Non-Weir customer {customerInfo.CustomerName} (Lite v1.0)");
                                logCallback?.Invoke($"⏭️ Skipped - non-Weir customer: {customerInfo.CustomerName}");
                                logCallback?.Invoke(""); // RESTORE LINE SPACING
                                progressManager?.IncrementSkippedEmails();
                                trackingData.SkippedEmailsCount++;
                                summary.SkippedEmails.Add(new SkippedEmailDetail
                                {
                                    Subject = email.Subject ?? "No Subject",
                                    Sender = email.From?.EmailAddress?.Address ?? "Unknown Sender",
                                    Reason = $"Non-Weir customer: {customerInfo.CustomerName} (Lite v1.0 - Weir only)",
                                    ReceivedDate = email.ReceivedDateTime?.DateTime ?? DateTime.Now
                                });
                                continue; // Skip this email completely
                            }
                            // STEP 3: Mark as processed
                            trackingData.ProcessedEmailIds.Add(emailKey);

                            // STEP 4: Try to extract business content
                            var emailDetail = await ProcessBusinessEmailWithRealTimeStats(
                                email, graphHandler, userEmail, progressManager, trackingData, processedCount, emailList.Count());

                            if (emailDetail != null && emailDetail.Orders.Any())
                            {
                                // SUCCESS: Business content found
                                summary.EmailDetails.Add(emailDetail);
                                summary.ProcessedEmails++;

                                // FIXED: Show BOTH headers and lines
                                var orderHeaders = emailDetail.Orders.Where(o => HasProperty(o, "PoNumber")).GroupBy(o => GetDynamicProperty(o, "PoNumber")).Count();
                                var orderLines = emailDetail.Orders.Count(o => HasProperty(o, "PoNumber"));

                                var quoteHeaders = emailDetail.Orders.Where(o => HasProperty(o, "RfqNumber") || HasProperty(o, "QuoteNumber")).GroupBy(o => GetDynamicProperty(o, "RfqNumber") + GetDynamicProperty(o, "QuoteNumber")).Count();
                                var quoteLines = emailDetail.Orders.Count(o => HasProperty(o, "RfqNumber") || HasProperty(o, "QuoteNumber"));

                                var revisionHeaders = emailDetail.Orders.Where(o => HasProperty(o, "CurrentRevision") || HasProperty(o, "NewRevision")).GroupBy(o => GetDynamicProperty(o, "ArtiCode")).Count();
                                var revisionLines = emailDetail.Orders.Count(o => HasProperty(o, "CurrentRevision") || HasProperty(o, "NewRevision"));

                                progressManager?.UpdateDebugConsole($"✅ Email {processedCount}/{emailList.Count()} processed: {emailDetail.Orders.Count} business items found");
                                logCallback?.Invoke($"   ✅ Found: {orderHeaders}H/{orderLines}L orders, {quoteHeaders}H/{quoteLines}L quotes, {revisionHeaders}H/{revisionLines}L revisions");
                                logCallback?.Invoke(""); // RESTORE LINE SPACING
                                LogWorkflow($"✅ Email processed successfully: {emailDetail.Subject}");
                            }
                            else
                            {
                                // FAILED: No business content found
                                LogWorkflow($"⏭️ Email skipped: No business content found");
                                logCallback?.Invoke($"   ⚠️ No business content found");
                                logCallback?.Invoke(""); // RESTORE LINE SPACING

                                // FIXED: This should increment S.Emails
                                Console.WriteLine($"🔥 DEBUG: Incrementing S.Emails for no business content");
                                progressManager?.IncrementSkippedEmails();
                                trackingData.SkippedEmailsCount++;

                                summary.SkippedEmails.Add(new SkippedEmailDetail
                                {
                                    Subject = email.Subject ?? "No Subject",
                                    Sender = email.From?.EmailAddress?.Address ?? "Unknown Sender",
                                    Reason = "No business content found",
                                    ReceivedDate = email.ReceivedDateTime?.DateTime ?? DateTime.Now
                                });
                            }
                        }
                        catch (Exception emailEx)
                        {
                            // ERROR: Processing failed
                            summary.FailedEmails++;
                            LogWorkflow($"❌ Error processing email: {emailEx.Message}");
                            logCallback?.Invoke($"❌ Error processing email: {emailEx.Message}");
                            logCallback?.Invoke(""); // RESTORE LINE SPACING
                            progressManager?.UpdateDebugConsole($"❌ Email {processedCount}/{emailList.Count()} failed: {emailEx.Message}");

                            // FIXED: Errors should increment S.Emails
                            Console.WriteLine($"🔥 DEBUG: Incrementing S.Emails for processing error");
                            progressManager?.IncrementSkippedEmails();
                            trackingData.SkippedEmailsCount++;

                            summary.SkippedEmails.Add(new SkippedEmailDetail
                            {
                                Subject = email.Subject ?? "No Subject",
                                Sender = email.From?.EmailAddress?.Address ?? "Unknown Sender",
                                Reason = $"Processing error: {emailEx.Message}",
                                ReceivedDate = email.ReceivedDateTime?.DateTime ?? DateTime.Now
                            });
                        }
                    }

                    // FIXED: Remove read-only property assignments
                    // The counts are calculated from the Lists automatically:
                    // summary.SkippedEmailsCount = summary.SkippedEmails.Count (read-only property)
                    // summary.DuplicateEmailsCount = summary.DuplicateEmails.Count (read-only property)

                    // DEBUG: Log final counts to verify (using correct property access)
                    LogWorkflow($"🔥 FINAL DEBUG COUNTS:");
                    LogWorkflow($"   T.Emails: {summary.TotalEmails}");
                    LogWorkflow($"   Processed: {summary.ProcessedEmails}");
                    LogWorkflow($"   S.Emails: {trackingData.SkippedEmailsCount} (Summary List: {summary.SkippedEmails.Count})");
                    LogWorkflow($"   D.Emails: {trackingData.DuplicateEmailsCount} (Summary List: {summary.DuplicateEmails.Count})");
                    LogWorkflow($"   Failed: {summary.FailedEmails}");

                    // FINAL CHECK: Update progress manager with correct counts
                    progressManager?.UpdateEmailStats(summary.TotalEmails, trackingData.SkippedEmailsCount, trackingData.DuplicateEmailsCount);

                    // Step 5: Calculate final summary totals (simplified version since methods don't exist)
                    progressManager?.UpdateProgress(80, 100, "Calculating summary totals...");
                    progressManager?.UpdateDebugConsole("📊 Calculating summary totals...");
                    logCallback?.Invoke("📊 Calculating final summary totals...");

                    // Simple summary calculation
                    summary.TotalOrdersExtracted = summary.EmailDetails.Sum(ed => ed.Orders.Count(o => HasProperty(o, "PoNumber")));
                    summary.TotalQuotesExtracted = summary.EmailDetails.Sum(ed => ed.Orders.Count(o => HasProperty(o, "RfqNumber") || HasProperty(o, "QuoteNumber")));
                    summary.TotalRevisionsExtracted = summary.EmailDetails.Sum(ed => ed.Orders.Count(o => HasProperty(o, "CurrentRevision") || HasProperty(o, "NewRevision")));

                    LogWorkflow($"📊 Summary calculation complete - {summary.TotalOrdersExtracted} orders, {summary.TotalQuotesExtracted} quotes, {summary.TotalRevisionsExtracted} revisions");
                    logCallback?.Invoke($"📊 Final totals: {summary.TotalOrdersExtracted} orders, {summary.TotalQuotesExtracted} quotes, {summary.TotalRevisionsExtracted} revisions");

                    // Step 6: RESTORE MISSING FUNCTIONALITY - Save to output tab and display in email import tab
                    progressManager?.UpdateProgress(90, 100, "Saving results to output...");
                    progressManager?.UpdateDebugConsole("💾 Saving results to output tab...");
                    logCallback?.Invoke("💾 Saving results to output and displaying in email import tab...");

                    // Generate detailed results content
                    var outputContent = GenerateDetailedResults(summary, trackingData);

                    // Save to output tab first
                    SaveToOutputTab(outputContent, $"Email Import Results - {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

                    // Then display in email import tab
                    DisplayInEmailImportTab(outputContent);

                    LogWorkflow("💾 Results saved to output tab and displayed in email import tab");
                    logCallback?.Invoke("💾 Results saved and displayed successfully");

                    // Step 7: Complete email import phase (but don't hide panel) and continue with injection
                    progressManager?.UpdateDebugConsole($"📧 Email import complete: {summary.ProcessedEmails} processed, {trackingData.SkippedEmailsCount} skipped, {trackingData.DuplicateEmailsCount} duplicates");
                    logCallback?.Invoke($"✅ Email import completed: {summary.ProcessedEmails} processed, {trackingData.SkippedEmailsCount} skipped, {trackingData.DuplicateEmailsCount} duplicates");
                    LogWorkflow($"=== EMAIL IMPORT COMPLETED ===");
                    LogWorkflow($"📊 Final Results: {summary.ProcessedEmails} processed, {summary.SkippedEmails.Count} skipped, {summary.FailedEmails} failed");
                    LogWorkflow($"📊 Real-time Statistics:");
                    LogWorkflow($"   📦 Total Order Headers: {trackingData.UniquePoNumbers.Count}");
                    LogWorkflow($"   📋 Total Order Lines: {trackingData.TotalOrderLines}");
                    LogWorkflow($"   💰 Total Quote Headers: {trackingData.UniqueRfqNumbers.Count}");
                    LogWorkflow($"   💱 Total Quote Lines: {trackingData.TotalQuoteLines}");
                    LogWorkflow($"   🔄 Total Revision Headers: {trackingData.UniqueRevisionArticles.Count}");
                    LogWorkflow($"   🔧 Total Revision Lines: {trackingData.TotalRevisionLines}");
                    LogWorkflow($"   ⏭️ Total Skipped Emails: {trackingData.SkippedEmailsCount}");
                    LogWorkflow($"   🔄 Total Duplicate Emails: {trackingData.DuplicateEmailsCount}");

                    // 🎯 Continue with MKG injection while keeping statistics visible
                    progressManager?.UpdateOperationPhase("MKG Injection Ready", 0, 100);
                    progressManager?.UpdateDebugConsole("📧 Email import completed successfully");
                    LogWorkflow($"=== EMAIL IMPORT COMPLETED - READY FOR INJECTION ===");
                    return summary;
                }
                finally
                {
                    // Clean up GraphHandler if it implements IDisposable
                    if (graphHandler is IDisposable disposable)
                    {
                        disposable.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                LogWorkflow($"❌ Critical error in email import: {ex.Message}");
                logCallback?.Invoke($"❌ Critical error in email import: {ex.Message}");
                summary.FailedEmails++;
                progressManager?.FailOperation($"Email import failed: {ex.Message}");

                // 🎯 NEW: Reset tab colors on error
                ResetAllTabColors();
                throw;
            }
        }

        #region Output Save and Display Methods - RESTORED FUNCTIONALITY

        /// <summary>
        /// Generate detailed results content for output tab
        /// </summary>
        private static string GenerateDetailedResults(EmailImportSummary summary, TrackingData trackingData)
        {
            var result = new
            {
                EmailImportResults = new
                {
                    GeneratedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    SummaryStatistics = new
                    {
                        TotalEmailsProcessed = summary.TotalEmails,
                        SuccessfullyProcessed = summary.ProcessedEmails,
                        SkippedEmails = summary.SkippedEmails.Count,
                        DuplicateEmails = summary.DuplicateEmails.Count,
                        FailedEmails = summary.FailedEmails
                    },
                    BusinessContentExtracted = new
                    {
                        OrderHeaders = trackingData.UniquePoNumbers.Count,
                        OrderLines = trackingData.TotalOrderLines,
                        QuoteHeaders = trackingData.UniqueRfqNumbers.Count,
                        QuoteLines = trackingData.TotalQuoteLines,
                        RevisionHeaders = trackingData.UniqueRevisionArticles.Count,
                        RevisionLines = trackingData.TotalRevisionLines
                    },
                    ProcessedEmails = summary.EmailDetails.Select(email => new
                    {
                        Subject = email.Subject,
                        Sender = email.Sender,
                        ClientDomain = email.ClientDomain,
                        ItemsFound = email.Orders.Count,
                        ReceivedDate = email.ReceivedDate.ToString("yyyy-MM-dd HH:mm")
                    }).ToArray(),
                    SkippedEmails = summary.SkippedEmails.Select(skipped => new
                    {
                        Subject = skipped.Subject,
                        Sender = skipped.Sender,
                        Reason = skipped.Reason,
                        ReceivedDate = skipped.ReceivedDate.ToString("yyyy-MM-dd HH:mm")
                    }).ToArray(),
                    DuplicateEmails = summary.DuplicateEmails.Select(duplicate => new
                    {
                        Subject = duplicate.Subject,
                        Sender = duplicate.Sender,
                        Reason = duplicate.Reason,
                        ReceivedDate = duplicate.ReceivedDate.ToString("yyyy-MM-dd HH:mm")
                    }).ToArray(),
                    IncrementalProcessingSteps = _processingLog.ToArray()
                }
            };

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            return JsonSerializer.Serialize(result, options);
        }
        private static void SaveToOutputTab(string content, string title)
        {
            try
            {
                // TODO: Replace this with your actual output tab save method
                // Example: OutputController.SaveToTab(content, title);
                LogWorkflow($"💾 Content saved to output tab: {title}");
            }
            catch (Exception ex)
            {
                LogWorkflow($"❌ Error saving to output tab: {ex.Message}");
            }
        }

        /// <summary>
        /// Display content in email import tab - PLACEHOLDER METHOD
        /// Implement this to call your existing email import tab display functionality
        /// </summary>
        private static void DisplayInEmailImportTab(string content)
        {
            try
            {
                // TODO: Replace this with your actual email import tab display method
                // Example: EmailImportController.DisplayResults(content);
                LogWorkflow("💾 Content displayed in email import tab");
            }
            catch (Exception ex)
            {
                LogWorkflow($"❌ Error displaying in email import tab: {ex.Message}");
            }
        }

        #endregion

        #region Real-time Statistics Increment Methods
        private static void IncrementSkippedEmails(EnhancedProgressManager progressManager, string reason)
        {
            progressManager?.IncrementSkippedEmails();
            LogWorkflow($"⏭️ S.Emails++ reason: {reason}");
        }

        private static void IncrementDuplicateEmails(EnhancedProgressManager progressManager, string duplicateInfo)
        {
            progressManager?.IncrementDuplicateEmails();
            LogWorkflow($"🔄 D.Emails++ duplicate: {duplicateInfo}");
        }

        private static void IncrementHeaderOrders(EnhancedProgressManager progressManager, string poNumber)
        {
            progressManager?.IncrementHeaderOrders();
            LogWorkflow($"📦 H.Orders++ for PO: {poNumber}");
        }

        private static void IncrementLineOrders(EnhancedProgressManager progressManager, string articleCode)
        {
            progressManager?.IncrementLineOrders();
            LogWorkflow($"📋 L.Orders++ for article: {articleCode}");
        }

        private static void IncrementHeaderQuotes(EnhancedProgressManager progressManager, string rfqNumber)
        {
            progressManager?.IncrementHeaderQuotes();
            LogWorkflow($"💰 H.Quotes++ for RFQ: {rfqNumber}");
        }

        private static void IncrementLineQuotes(EnhancedProgressManager progressManager, string quoteItem)
        {
            progressManager?.IncrementLineQuotes();
            LogWorkflow($"💱 L.Quotes++ for quote: {quoteItem}");
        }

        private static void IncrementHeaderRevisions(EnhancedProgressManager progressManager, string articleCode)
        {
            progressManager?.IncrementHeaderRevisions();
            LogWorkflow($"🔄 H.Revisions++ for article: {articleCode}");
        }

        private static void IncrementLineRevisions(EnhancedProgressManager progressManager, string revisionDetail)
        {
            progressManager?.IncrementLineRevisions();
            LogWorkflow($"🔧 L.Revisions++ for revision: {revisionDetail}");
        }

        #endregion

        #region Helper Classes

        /// <summary>
        /// Helper class to track unique items across all emails for real-time header counting
        /// </summary>
        private class TrackingData
        {
            public HashSet<string> UniquePoNumbers { get; } = new HashSet<string>();
            public HashSet<string> UniqueRfqNumbers { get; } = new HashSet<string>();
            public HashSet<string> UniqueRevisionArticles { get; } = new HashSet<string>();
            public int TotalOrderLines { get; set; } = 0;
            public int TotalQuoteLines { get; set; } = 0;
            public int TotalRevisionLines { get; set; } = 0;
            public HashSet<string> ProcessedEmailIds { get; } = new HashSet<string>();
            public int SkippedEmailsCount { get; set; } = 0;
            public int DuplicateEmailsCount { get; set; } = 0;
        }

        #endregion

        #region Conversion Methods for MKG Injection

        /// <summary>
        /// Convert email summary to order lines for MKG injection
        /// </summary>
        public static List<OrderLine> ConvertEmailSummaryToOrderLines(EmailImportSummary summary)
        {
            var orderLines = new List<OrderLine>();

            foreach (var emailDetail in summary.EmailDetails)
            {
                foreach (dynamic order in emailDetail.Orders)
                {
                    // Only add items that have PoNumber (are actual orders)
                    if (HasProperty(order, "PoNumber"))
                    {
                        orderLines.Add(new OrderLine
                        {
                            PoNumber = GetDynamicProperty(order, "PoNumber"),
                            ArtiCode = GetDynamicProperty(order, "ArtiCode"),
                            Quantity = GetDynamicProperty(order, "Quantity"),
                            Description = GetDynamicProperty(order, "Description"),
                            Unit = GetDynamicProperty(order, "Unit"),
                            UnitPrice = GetDynamicProperty(order, "UnitPrice"),
                            TotalPrice = GetDynamicProperty(order, "TotalPrice"),
                            DeliveryDate = GetDynamicProperty(order, "DeliveryDate"),
                            EmailDomain = emailDetail.ClientDomain,
                            ExtractionMethod = GetDynamicProperty(order, "ExtractionMethod"),
                            LineNumber = GetDynamicProperty(order, "LineNumber"),
                            Notes = GetDynamicProperty(order, "Notes"),
                            Priority = GetDynamicProperty(order, "Priority") ?? "NORMAL"
                        });
                    }
                }
            }

            return orderLines;
        }

        /// <summary>
        /// Convert email summary to quote lines for MKG injection
        /// </summary>
        public static List<QuoteLine> ConvertEmailSummaryToQuoteLines(EmailImportSummary summary)
        {
            var quoteLines = new List<QuoteLine>();

            foreach (var emailDetail in summary.EmailDetails)
            {
                foreach (dynamic item in emailDetail.Orders)
                {
                    if (HasProperty(item, "RfqNumber"))
                    {
                        quoteLines.Add(new QuoteLine
                        {
                            RfqNumber = GetDynamicProperty(item, "RfqNumber"),
                            ArtiCode = GetDynamicProperty(item, "ArtiCode"),
                            Description = GetDynamicProperty(item, "Description"),
                            Quantity = GetDynamicProperty(item, "Quantity"),
                            Unit = GetDynamicProperty(item, "Unit"),
                            QuotedPrice = GetDynamicProperty(item, "QuotedPrice"),
                            ValidUntil = GetDynamicProperty(item, "ValidUntil"),
                            EmailDomain = emailDetail.ClientDomain,
                            ExtractionMethod = GetDynamicProperty(item, "ExtractionMethod"),
                            Priority = GetDynamicProperty(item, "Priority") ?? "NORMAL"
                            // ✅ Notes property removed
                        });
                    }
                }
            }

            return quoteLines;
        }

        /// <summary>
        /// Convert email summary to revision lines for MKG injection
        /// </summary>
        public static List<RevisionLine> ConvertEmailSummaryToRevisionLines(EmailImportSummary summary)
        {
            var revisionLines = new List<RevisionLine>();

            foreach (var emailDetail in summary.EmailDetails)
            {
                foreach (dynamic item in emailDetail.Orders)
                {
                    // Only add items that have CurrentRevision or NewRevision (are actual revisions)
                    if (HasProperty(item, "CurrentRevision") || HasProperty(item, "NewRevision"))
                    {
                        revisionLines.Add(new RevisionLine
                        {
                            ArtiCode = GetDynamicProperty(item, "ArtiCode"),
                            CurrentRevision = GetDynamicProperty(item, "CurrentRevision"),
                            NewRevision = GetDynamicProperty(item, "NewRevision"),
                            RevisionReason = GetDynamicProperty(item, "RevisionReason"),
                            Description = GetDynamicProperty(item, "Description"),
                            EmailDomain = emailDetail.ClientDomain,
                            ExtractionMethod = GetDynamicProperty(item, "ExtractionMethod"),
                            DrawingNumber = GetDynamicProperty(item, "DrawingNumber"),
                            TechnicalChanges = GetDynamicProperty(item, "TechnicalChanges"),
                            RevisionDate = emailDetail.ReceivedDate.ToString("yyyy-MM-dd"),
                            FieldChanged = GetDynamicProperty(item, "FieldChanged"),
                            OldValue = GetDynamicProperty(item, "OldValue"),
                            NewValue = GetDynamicProperty(item, "NewValue")
                        });
                    }
                }
            }

            return revisionLines;
        }

        #endregion
        #region Email Processing Methods with Live Statistics

        /// <summary>
        /// Process business email with complete extraction - Enhanced with real-time statistics tracking
        /// </summary>
        private static async Task<EmailDetail> ProcessBusinessEmailWithRealTimeStats(
      Message email,
      GraphHandler graphHandler,
      string userEmail,
      EnhancedProgressManager progressManager,
      TrackingData trackingData,
      int currentEmailNumber,  // 🎯 NEW PARAMETER
      int totalEmails)         // 🎯 NEW PARAMETER
        {
            try
            {
                // Step 1: Basic email detail setup
                var emailDetail = new EmailDetail
                {
                    Subject = email.Subject ?? "No Subject",
                    Sender = email.From?.EmailAddress?.Address ?? "Unknown Sender",
                    ReceivedDate = email.ReceivedDateTime?.DateTime ?? DateTime.Now,
                    ClientDomain = ExtractDomain(email.From?.EmailAddress?.Address ?? ""),
                    Body = email.Body?.Content ?? ""
                };

                LogWorkflow($"🔄 Processing email {currentEmailNumber}/{totalEmails}: '{emailDetail.Subject}' from {emailDetail.Sender}");

                // Step 2: Get email body and attachments (FIXED)
                var attachmentResponse = await graphHandler.GetAttachementFromEmailWithId(userEmail, email.Id);
                emailDetail.Attachments = attachmentResponse?.Value?.Select(a => new Mkg_Elcotec_Automation.Models.EmailModels.AttachmentInfo
                {
                    Name = a.Name,
                    Extension = System.IO.Path.GetExtension(a.Name ?? ""),
                    Size = a.Size ?? 0,
                    ContentType = a.ContentType ?? ""
                }).ToList() ?? new List<Mkg_Elcotec_Automation.Models.EmailModels.AttachmentInfo>();

                // Step 3: Extract all business content with real-time statistics
                await ExtractAllBusinessContentWithRealTimeStats(emailDetail, emailDetail.Body,
                    attachmentResponse, progressManager, trackingData);

                // Step 4: Check if any business content was found
                if (emailDetail.Orders.Any())
                {
                    LogWorkflow($"✅ Email {currentEmailNumber}/{totalEmails} processed: {emailDetail.Orders.Count} business items extracted");
                    return emailDetail;
                }
                else
                {
                    LogWorkflow($"⏭️ Email {currentEmailNumber}/{totalEmails} skipped: No business content found");
                    return null;
                }
            }
            catch (Exception ex)
            {
                LogWorkflow($"❌ Error processing email {currentEmailNumber}/{totalEmails} '{email.Subject}': {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Extract all business content (Orders + Quotes + Revisions) - Enhanced with real-time statistics
        /// </summary>
        private static async Task ExtractAllBusinessContentWithRealTimeStats(
    EmailDetail emailDetail,
    string emailBody,
    AttachmentCollectionResponse attachmentResponse,  // CHANGED TYPE
    EnhancedProgressManager progressManager,
    TrackingData trackingData)
        {
            try
            {
                LogWorkflow("🎯 Starting comprehensive business content extraction with real-time stats...");

                // Extract Orders with real-time statistics
                await ExtractOrdersWithRealTimeStats(emailDetail, emailBody, attachmentResponse, progressManager, trackingData);

                // Extract Quotes with real-time statistics  
                await ExtractQuotesWithRealTimeStats(emailDetail, emailBody, attachmentResponse, progressManager, trackingData);

                // Extract Revisions with real-time statistics
                await ExtractRevisionsWithRealTimeStats(emailDetail, emailBody, attachmentResponse, progressManager, trackingData);

                LogWorkflow($"✅ Business content extraction complete: {emailDetail.Orders.Count} total items");
            }
            catch (Exception ex)
            {
                LogWorkflow($"❌ Error in business content extraction: {ex.Message}");
            }
        }

        /// <summary>
        /// Extract orders using working pattern - Enhanced with real-time statistics
        /// </summary>
        private static async Task ExtractOrdersWithRealTimeStats(
            EmailDetail emailDetail,
            string emailBody,
            AttachmentCollectionResponse attachments,
            EnhancedProgressManager progressManager,
            TrackingData trackingData)
        {
            try
            {
                LogWorkflow("📦 Extracting orders using WORKING PATTERN with real-time stats");
                var extractedOrders = new List<OrderLine>();
                var uniquePoNumbersInEmail = new HashSet<string>();

                // Step 1: Try HTML parser first using optimal parsing method
                if (!string.IsNullOrEmpty(emailBody))
                {
                    try
                    {
                        // Determine optimal parsing method based on domain - simplified approach
                        var parsingMethod = emailDetail.Sender.Contains("weir") ?
                            "HtmlTableParser" : "HtmlGenericParser";

                        var htmlOrders = await OrderLogicService.ExtractOrders(emailBody, emailDetail.ClientDomain, emailDetail.Subject, attachments);

                        if (htmlOrders.Any())
                        {
                            extractedOrders.AddRange(htmlOrders);
                            LogWorkflow($"📦 HTML extraction: Found {htmlOrders.Count} orders");
                        }
                    }
                    catch (Exception htmlEx)
                    {
                        LogWorkflow($"⚠️ HTML order extraction failed: {htmlEx.Message}");
                    }
                }

                // Step 2: Subject parsing as fallback
                if (!extractedOrders.Any())
                {
                    try
                    {
                        var subjectOrders = await OrderLogicService.ExtractOrders(emailDetail.Subject, emailDetail.ClientDomain, emailDetail.Subject, null);

                        if (subjectOrders.Any())
                        {
                            extractedOrders.AddRange(subjectOrders);
                            LogWorkflow($"📦 Subject extraction: Found {subjectOrders.Count} orders");
                        }
                    }
                    catch (Exception subjectEx)
                    {
                        LogWorkflow($"⚠️ Subject order extraction failed: {subjectEx.Message}");
                    }
                }

                // Step 3: Process extracted orders with real-time statistics
                foreach (var order in extractedOrders)
                {
                    // Ensure proper email domain assignment
                    order.EmailDomain = emailDetail.ClientDomain;

                    // Add to email detail
                    emailDetail.Orders.Add(order);

                    // Real-time statistics tracking
                    if (!string.IsNullOrEmpty(order.PoNumber))
                    {
                        // Track unique PO numbers for header counting
                        if (uniquePoNumbersInEmail.Add(order.PoNumber))
                        {
                            if (trackingData.UniquePoNumbers.Add(order.PoNumber))
                            {
                                // New unique PO across all emails - increment header
                                IncrementHeaderOrders(progressManager, order.PoNumber);
                            }
                        }

                        // Always increment line orders for each order line
                        trackingData.TotalOrderLines++;
                        IncrementLineOrders(progressManager, order.ArtiCode ?? "UNKNOWN");
                    }
                }

                LogWorkflow($"📦 Orders extraction completed: {extractedOrders.Count} orders found");
            }
            catch (Exception ex)
            {
                LogWorkflow($"❌ Error in orders extraction: {ex.Message}");
            }
        }

        /// <summary>
        /// Extract quotes with real-time statistics tracking
        /// </summary>
        private static async Task ExtractQuotesWithRealTimeStats(
            EmailDetail emailDetail,
            string emailBody,
            AttachmentCollectionResponse attachments,
            EnhancedProgressManager progressManager,
            TrackingData trackingData)
        {
            try
            {
                LogWorkflow("💰 Extracting quotes with real-time stats");
                var extractedQuotes = new List<QuoteLine>();
                var uniqueRfqNumbersInEmail = new HashSet<string>();

                // Extract quotes using QuoteLogicService
                if (!string.IsNullOrEmpty(emailBody))
                {
                    try
                    {
                        var htmlQuotes = await QuoteLogicService.ExtractQuotesSafe(emailBody, emailDetail.ClientDomain, emailDetail.Subject, attachments);

                        if (htmlQuotes.Any())
                        {
                            extractedQuotes.AddRange(htmlQuotes);
                            LogWorkflow($"💰 Quote extraction: Found {htmlQuotes.Count} quotes");
                        }
                    }
                    catch (Exception quoteEx)
                    {
                        LogWorkflow($"⚠️ Quote extraction failed: {quoteEx.Message}");
                    }
                }

                // Process extracted quotes with real-time statistics
                foreach (var quote in extractedQuotes)
                {
                    quote.EmailDomain = emailDetail.ClientDomain;
                    emailDetail.Orders.Add(quote);

                    // Real-time statistics tracking
                    if (!string.IsNullOrEmpty(quote.RfqNumber))
                    {
                        // Track unique RFQ numbers for header counting
                        if (uniqueRfqNumbersInEmail.Add(quote.RfqNumber))
                        {
                            if (trackingData.UniqueRfqNumbers.Add(quote.RfqNumber))
                            {
                                IncrementHeaderQuotes(progressManager, quote.RfqNumber);
                            }
                        }

                        // Always increment line quotes
                        trackingData.TotalQuoteLines++;
                        IncrementLineQuotes(progressManager, quote.ArtiCode ?? "UNKNOWN");
                    }
                }

                LogWorkflow($"💰 Quotes extraction completed: {extractedQuotes.Count} quotes found");
            }
            catch (Exception ex)
            {
                LogWorkflow($"❌ Error in quotes extraction: {ex.Message}");
            }
        }

        /// <summary>
        /// Extract revisions with real-time statistics tracking
        /// </summary>
        private static async Task ExtractRevisionsWithRealTimeStats(
            EmailDetail emailDetail,
            string emailBody,
            AttachmentCollectionResponse attachments,
            EnhancedProgressManager progressManager,
            TrackingData trackingData)
        {
            try
            {
                LogWorkflow("🔄 Extracting revisions with real-time stats");
                var extractedRevisions = new List<RevisionLine>();
                var uniqueRevisionArticlesInEmail = new HashSet<string>();

                // Extract revisions using RevisionLogicService
                if (!string.IsNullOrEmpty(emailBody))
                {
                    try
                    {
                        var htmlRevisions = await RevisionLogicService.ExtractRevisionsSafe(emailBody, emailDetail.ClientDomain, emailDetail.Subject, attachments);

                        if (htmlRevisions.Any())
                        {
                            extractedRevisions.AddRange(htmlRevisions);
                            LogWorkflow($"🔄 Revision extraction: Found {htmlRevisions.Count} revisions");
                        }
                    }
                    catch (Exception revisionEx)
                    {
                        LogWorkflow($"⚠️ Revision extraction failed: {revisionEx.Message}");
                    }
                }

                // Process extracted revisions with real-time statistics
                foreach (var revision in extractedRevisions)
                {
                    revision.EmailDomain = emailDetail.ClientDomain;
                    emailDetail.Orders.Add(revision);

                    // Real-time statistics tracking
                    if (!string.IsNullOrEmpty(revision.ArtiCode))
                    {
                        // Track unique article codes for header counting
                        if (uniqueRevisionArticlesInEmail.Add(revision.ArtiCode))
                        {
                            if (trackingData.UniqueRevisionArticles.Add(revision.ArtiCode))
                            {
                                IncrementHeaderRevisions(progressManager, revision.ArtiCode);
                            }
                        }

                        // Always increment line revisions
                        trackingData.TotalRevisionLines++;
                        IncrementLineRevisions(progressManager, $"{revision.ArtiCode}: {revision.CurrentRevision}→{revision.NewRevision}");
                    }
                }

                LogWorkflow($"🔄 Revisions extraction completed: {extractedRevisions.Count} revisions found");
            }
            catch (Exception ex)
            {
                LogWorkflow($"❌ Error in revisions extraction: {ex.Message}");
            }
        }

        #endregion

        #region Utility Methods

        /// <summary>
        /// Extract domain from email address
        /// </summary>
        private static string ExtractDomain(string email)
        {
            if (string.IsNullOrEmpty(email) || !email.Contains("@"))
                return "";

            return email.Substring(email.IndexOf("@") + 1);
        }

        /// <summary>
        /// Check if dynamic object has property - FIXED implementation
        /// </summary>
        private static bool HasProperty(dynamic obj, string propertyName)
        {
            try
            {
                if (obj == null) return false;
                var objType = obj.GetType();
                return objType.GetProperty(propertyName) != null;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Get dynamic property value safely
        /// </summary>
        private static string GetDynamicProperty(dynamic obj, string propertyName)
        {
            try
            {
                var objType = obj.GetType();
                var property = objType.GetProperty(propertyName);
                return property?.GetValue(obj)?.ToString() ?? "";
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Update progress with current statistics summary
        /// </summary>
        private static void UpdateProgressWithStatistics(EnhancedProgressManager progressManager, TrackingData trackingData)
        {
            if (progressManager != null)
            {
                progressManager.UpdateDebugConsole($"Current Stats - Orders={trackingData.UniquePoNumbers.Count}, Quotes={trackingData.UniqueRfqNumbers.Count}, Revisions={trackingData.UniqueRevisionArticles.Count}");
            }
        }

        #endregion

        #region Email Processing Helper Methods

        /// <summary>
        /// Generate unique key for email to detect duplicates
        /// </summary>
        private static string GenerateEmailKey(Message email)
        {
            try
            {
                var sender = email.From?.EmailAddress?.Address ?? "unknown";
                var subject = email.Subject ?? "no-subject";
                var bodyPreview = email.BodyPreview ?? "";

                // DEBUG: Always log what we're processing
                LogWorkflow($"🔍 DUPLICATE CHECK: From={sender}");
                LogWorkflow($"🔍 DUPLICATE CHECK: Subject='{subject}'");
                LogWorkflow($"🔍 DUPLICATE CHECK: BodyPreview='{bodyPreview?.Substring(0, Math.Min(bodyPreview.Length, 100))}'");

                // ENHANCED: Extract PO number for Weir duplicate detection
                var senderDomain = ExtractDomain(sender);
                if (senderDomain.Contains("weir", StringComparison.OrdinalIgnoreCase))
                {
                    LogWorkflow($"🔍 WEIR DOMAIN DETECTED: {senderDomain}");

                    // Look for PO patterns in subject and body
                    var searchText = subject + " " + bodyPreview;
                    var poPattern = @"(?:po|order)[:\s#-]*(\d{8,12})";
                    var match = Regex.Match(searchText, poPattern, RegexOptions.IgnoreCase);

                    LogWorkflow($"🔍 SEARCHING: '{searchText}' with pattern '{poPattern}'");

                    if (match.Success)
                    {
                        var extractedPO = match.Groups[1].Value;
                        var duplicateKey = $"WEIR_PO_{extractedPO}";
                        LogWorkflow($"🎯 PO EXTRACTED: '{extractedPO}' → KEY: '{duplicateKey}'");
                        return duplicateKey;
                    }
                    else
                    {
                        LogWorkflow($"❌ NO PO MATCH FOUND in Weir email");
                    }
                }
                else
                {
                    LogWorkflow($"🔍 NON-WEIR DOMAIN: {senderDomain}");
                }

                // Fallback to original logic for non-Weir or emails without PO numbers
                var combined = $"{sender}|{subject}|{bodyPreview}".ToLower();
                var fallbackKey = combined.GetHashCode().ToString();
                LogWorkflow($"🔄 FALLBACK KEY: {fallbackKey}");
                return fallbackKey;
            }
            catch
            {
                return Guid.NewGuid().ToString(); // Fallback to unique ID
            }
        }
        #endregion
    }
}