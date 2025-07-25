using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using Mkg_Elcotec_Automation.Services;
using System.Configuration;
using Mkg_Elcotec_Automation.Models;
using System.Text.Json;
using System.Net.Http;

namespace Mkg_Elcotec_Automation.Controllers
{
    /// <summary>
    /// MKG Quote Controller with Enhanced Debug Logging - Fixed Ambiguous References
    /// </summary>
    public class MkgQuoteController : MkgController
    {
        #region Configuration Properties
        private int AdministrationNumber => int.Parse(ConfigurationManager.AppSettings["MkgApi:AdministrationNumber"] ?? "1");
        private int DebtorNumber => int.Parse(ConfigurationManager.AppSettings["MkgApi:DebtorNumber"] ?? "30010");
        private int RelationNumber => int.Parse(ConfigurationManager.AppSettings["MkgApi:RelationNumber"] ?? "2");
        #endregion

        #region Debug Logging
        private void LogDebug(string message)
        {
            DebugLogger.LogDebug(message, "MkgQuoteController");
        }

        private void LogSection(string sectionName)
        {
            DebugLogger.LogDebug("Starting quote processing section", "MkgQuoteController");
        }

        private void LogError(string message, Exception ex = null)
        {
            DebugLogger.LogError("Error message", ex, "MkgQuoteController");  // Use 'ex', not 'exception'
        }
        #endregion

        #region Public Methods

        /// <summary>
        /// Enhanced InjectQuotesAsync with comprehensive debug logging
        /// </summary>
        public async Task<MkgQuoteInjectionSummary> InjectQuotesAsync(List<QuoteLine> quoteLines,
                                                                    IProgress<string> progress = null,
                                                                    EnhancedProgressManager progressManager = null)
        {
            LogSection("MKG QUOTE INJECTION STARTED");

            var summary = new MkgQuoteInjectionSummary
            {
                StartTime = DateTime.Now,
                TotalQuotes = 0
            };

            try
            {
                LogDebug($"🔄 Starting MKG quote injection for {quoteLines.Count} QuoteLines");
                LogDebug($"🔄 Debug log file: {DebugLogger.GetCurrentLogFile()}");

                progress?.Report($"Starting injection of {quoteLines.Count} QuoteLines...");

                if (!quoteLines.Any())
                {
                    LogDebug("⚠️ No QuoteLines provided for injection");
                    summary.EndTime = DateTime.Now;
                    summary.ProcessingTime = summary.EndTime - summary.StartTime;
                    return summary;
                }

                // Step 1: Filter valid quote lines and group by RFQ Number
                LogSection("FILTERING AND GROUPING QUOTE LINES");
                var validQuoteLines = FilterValidQuoteLines(quoteLines);
                var quoteGroups = validQuoteLines.GroupBy(ql => ql.RfqNumber).ToList();

                LogDebug($"💰 Original QuoteLines: {quoteLines.Count}");
                LogDebug($"💰 Valid QuoteLines after filtering: {validQuoteLines.Count}");
                LogDebug($"💰 Grouped into {quoteGroups.Count} Quote Headers");

                // Log details of what will be processed
                foreach (var group in quoteGroups)
                {
                    LogDebug($"💰 RFQ Group: '{group.Key}' contains {group.Count()} lines:");
                    foreach (var line in group)
                    {
                        LogDebug($"   └── Article: {line.ArtiCode} | Desc: {line.Description} | Price: €{line.QuotedPrice}");
                    }
                }

                summary.TotalQuotes = quoteGroups.Count;

                foreach (var quoteGroup in quoteGroups)
                {
                    var rfqNumber = quoteGroup.Key;
                    var linesInQuote = quoteGroup.ToList();

                    try
                    {
                        LogSection($"PROCESSING RFQ: {rfqNumber}");
                        LogDebug($"🔄 Processing RFQ: {rfqNumber} with {linesInQuote.Count} QuoteLines");

                        progress?.Report($"Checking for duplicates: RFQ {rfqNumber}...");

                        // ✅ Check for existing RFQ (BLOCKING CHECK)
                        LogDebug($"🔍 RFQ CHECK: Starting duplicate check for RFQ: {rfqNumber}");
                        var rfqExists = await CheckForExistingRfq(rfqNumber);

                        if (rfqExists)
                        {
                            LogSection("RFQ DUPLICATE DETECTED - BLOCKING INJECTION");
                            LogDebug($"⚠️ RFQ DUPLICATE: RFQ '{rfqNumber}' already exists in MKG system");
                            LogDebug($"⚠️ BLOCKING: Skipping injection of all {linesInQuote.Count} lines in this RFQ");
                            progress?.Report($"⚠️ Skipped duplicate RFQ: {rfqNumber}");

                            // Mark all lines as skipped duplicates
                            foreach (var line in linesInQuote)
                            {
                                summary.QuoteResults.Add(new MkgQuoteResult
                                {
                                    ArtiCode = line.ArtiCode,
                                    RfqNumber = line.RfqNumber,
                                    Success = false,
                                    ErrorMessage = $"RFQ '{rfqNumber}' already exists - MKG Duplicate",
                                    ProcessedAt = DateTime.Now,
                                    HttpStatusCode = "MKG_DUPLICATE_SKIPPED"
                                });

                                // ✅ ADD THIS LINE - INCREMENT DUPLICATE ERRORS for live statistics
                                progressManager?.IncrementDuplicateErrors();
                            }

                            summary.DuplicatesFiltered++;
                            continue; // Skip to next RFQ
                        }
                        else
                        {
                            LogDebug($"✅ RFQ CHECK: RFQ '{rfqNumber}' is unique, proceeding with injection");
                        }

                        // ✅ Create Quote Header in MKG
                        LogSection($"CREATING MKG QUOTE HEADER FOR RFQ: {rfqNumber}");
                        LogDebug($"🔄 Creating MKG Quote Header for RFQ: {rfqNumber} with {linesInQuote.Count} lines");
                        progress?.Report($"Creating MKG Quote for RFQ: {rfqNumber} ({linesInQuote.Count} lines)...");

                        var quoteHeaderResult = await CreateMkgQuoteHeader(rfqNumber, linesInQuote.First());

                        if (!quoteHeaderResult.Success)
                        {
                            LogDebug($"❌ HEADER CREATION FAILED: RFQ '{rfqNumber}' - {quoteHeaderResult.ErrorMessage}");
                            progress?.Report($"❌ Failed to create quote header for RFQ: {rfqNumber}");

                            // Mark all lines as failed
                            foreach (var line in linesInQuote)
                            {
                                summary.QuoteResults.Add(new MkgQuoteResult
                                {
                                    ArtiCode = line.ArtiCode,
                                    RfqNumber = line.RfqNumber,
                                    Success = false,
                                    ErrorMessage = $"Quote header creation failed: {quoteHeaderResult.ErrorMessage}",
                                    ProcessedAt = DateTime.Now,
                                    HttpStatusCode = quoteHeaderResult.HttpStatusCode
                                });
                                summary.FailedInjections++;
                            }
                            continue;
                        }
                        var mkgQuoteNumber = quoteHeaderResult.MkgQuoteNumber;
                        if (quoteHeaderResult.Success)
                        {
                            LogDebug($"✅ Created NEW MKG Quote Header: {mkgQuoteNumber} for RFQ: {rfqNumber}");
                            progress?.Report($"✅ Created MKG Quote: {mkgQuoteNumber}, now adding {linesInQuote.Count} lines...");
                        }
                        // ✅ Quote header created successfully
                        LogDebug($"🔄 QUOTE LINES: Now injecting {linesInQuote.Count} quote lines...");

                        // Step 4: Inject individual quote lines
                        LogSection($"INJECTING QUOTE LINES FOR MKG QUOTE: {mkgQuoteNumber}");
                        foreach (var quoteLine in linesInQuote)
                        {
                            try
                            {
                                var quoteLineResult = await InjectSingleQuoteLine(quoteLine, mkgQuoteNumber);

                                summary.QuoteResults.Add(new MkgQuoteResult
                                {
                                    ArtiCode = quoteLine.ArtiCode,
                                    RfqNumber = quoteLine.RfqNumber,
                                    MkgQuoteId = mkgQuoteNumber,
                                    Success = quoteLineResult.Success,
                                    ErrorMessage = quoteLineResult.ErrorMessage,
                                    ProcessedAt = DateTime.Now,
                                    HttpStatusCode = quoteLineResult.HttpStatusCode
                                });

                                if (quoteLineResult.Success)
                                {
                                    summary.SuccessfulInjections++;

                                    // ❌ REMOVED: progressManager?.IncrementLineQuotes(); // This was wrong during injection

                                    LogDebug($"✅ QuoteLine {quoteLine.ArtiCode} → MKG Quote {mkgQuoteNumber}");
                                }
                                else
                                {
                                    summary.FailedInjections++;

                                    // ✅ KEEP: This is correct during injection - increment error counters
                                    progressManager?.IncrementInjectionErrors();

                                    LogDebug($"❌ QuoteLine {quoteLine.ArtiCode} failed: {quoteLineResult.ErrorMessage}");
                                }
                            }
                            catch (Exception lineEx)
                            {
                                summary.FailedInjections++;
                                progressManager?.IncrementInjectionErrors();
                                LogDebug($"❌ QuoteLine {quoteLine.ArtiCode} exception: {lineEx.Message}");
                            }
                        }


                        LogDebug($"✅ RFQ COMPLETED: RFQ '{rfqNumber}' injection completed");
                    }
                    catch (Exception groupEx)
                    {
                        LogError($"RFQ '{rfqNumber}' threw exception", groupEx);

                        // Mark all lines in this group as failed
                        foreach (var line in linesInQuote)
                        {
                            summary.QuoteResults.Add(new MkgQuoteResult
                            {
                                ArtiCode = line.ArtiCode,
                                RfqNumber = line.RfqNumber,
                                Success = false,
                                ErrorMessage = groupEx.Message,
                                ProcessedAt = DateTime.Now,
                                HttpStatusCode = "GROUP_ERROR"
                            });
                            summary.FailedInjections++;
                        }
                    }
                }

                // ✅ Finalize summary
                summary.EndTime = DateTime.Now;
                summary.ProcessingTime = summary.EndTime - summary.StartTime;

                LogSection("MKG QUOTE INJECTION COMPLETED");
                LogDebug($"🏁 SUMMARY: Quote injection completed");
                LogDebug($"🏁 SUMMARY: Successful injections: {summary.SuccessfulInjections}");
                LogDebug($"🏁 SUMMARY: Failed injections: {summary.FailedInjections}");
                LogDebug($"🏁 SUMMARY: Duplicates filtered: {summary.DuplicatesFiltered}");
                LogDebug($"🏁 SUMMARY: Processing time: {summary.ProcessingTime.TotalSeconds:F1} seconds");

                return summary;
            }
            catch (Exception ex)
            {
                LogError("Quote injection process failed", ex);

                summary.Errors.Add($"Quote injection process error: {ex.Message}");
                summary.EndTime = DateTime.Now;
                summary.ProcessingTime = summary.EndTime - summary.StartTime;

                throw;
            }
        }

        #endregion

        #region Helper Methods

        private List<QuoteLine> FilterValidQuoteLines(List<QuoteLine> quoteLines)
        {
            LogDebug($"🔄 FILTERING: Starting filter of {quoteLines.Count} quote lines");

            var validQuoteLines = quoteLines.Where(ql =>
                !string.IsNullOrEmpty(ql.ArtiCode) &&
                !string.IsNullOrEmpty(ql.RfqNumber) &&
                ql.ArtiCode != "UNKNOWN-QUOTE" &&
                ql.ArtiCode.Length > 2
            ).ToList();

            var filteredCount = quoteLines.Count - validQuoteLines.Count;
            LogDebug($"🔄 FILTERING: Filtered out {filteredCount} invalid quote lines");
            LogDebug($"🔄 FILTERING: {validQuoteLines.Count} valid quote lines remaining");

            if (filteredCount > 0)
            {
                LogDebug($"🔄 FILTERING: Invalid quote lines filtered:");
                foreach (var invalid in quoteLines.Except(validQuoteLines))
                {
                    LogDebug($"   ❌ Filtered: '{invalid.ArtiCode}' (RFQ: {invalid.RfqNumber}) - Reason: Invalid format");
                }
            }

            return validQuoteLines;
        }

        private async Task<bool> CheckForExistingRfq(string rfqNumber)
        {
            try
            {
                LogDebug($"🔍 RFQ CHECK: Checking for existing RFQ: '{rfqNumber}'");

                if (string.IsNullOrEmpty(rfqNumber))
                {
                    LogDebug($"⚠️ RFQ CHECK: Empty RFQ number provided, skipping check");
                    return false;
                }

                // Search for existing quotes using RFQ reference fields
                var searchQuery = $"vofh_ref_extern={rfqNumber}";
                var encodedQuery = Uri.EscapeDataString(searchQuery);
                var endpoint = $"Documents/vofh/{AdministrationNumber}?Filter={encodedQuery}&FieldList=vofh_num,vofh_ref_extern&NumRows=5";

                LogDebug($"🔍 RFQ CHECK: API endpoint: {endpoint}");

                var responseBody = await _mkgApiClient.GetAsync(endpoint);

                LogDebug($"🔍 RFQ CHECK: Response length: {responseBody?.Length ?? 0}");
                LogDebug($"🔍 RFQ CHECK: Response preview: {responseBody?.Substring(0, Math.Min(responseBody?.Length ?? 0, 200))}...");

                if (!string.IsNullOrEmpty(responseBody) && !responseBody.Contains("error"))
                {
                    if (responseBody.Contains("vofh_num") && !responseBody.Contains("\"data\":[]"))
                    {
                        LogDebug($"✅ RFQ DUPLICATE FOUND: RFQ '{rfqNumber}' already exists in MKG system");
                        return true;
                    }
                }

                LogDebug($"✅ RFQ CHECK: No duplicate found for RFQ: '{rfqNumber}'");
                return false;
            }
            catch (Exception ex)
            {
                LogError($"RFQ duplicate check failed for '{rfqNumber}'", ex);
                return false;
            }
        }

        /// <summary>
        /// Create MKG Quote Header for a specific RFQ
        /// </summary>
        private async Task<MkgQuoteHeaderResult> CreateMkgQuoteHeader(string rfqNumber, QuoteLine firstQuote)
        {
            try
            {
                LogDebug($"🔄 Creating MKG Quote Header for RFQ: {rfqNumber}");

                var customerInfo = await FindCustomerByEmailDomain(firstQuote.EmailDomain);

                var quoteHeaderData = new
                {
                    request = new
                    {
                        InputData = new
                        {
                            vofh = new[]
                            {
                        new
                        {
                            admi_num = customerInfo?.AdministrationNumber ?? "1",
                            debi_num = customerInfo?.DebtorNumber ?? "99999",
                            rela_num = customerInfo?.RelationNumber ?? "99999",
                            vofh_referentie = rfqNumber,
                            vofh_omschrijving = $"Quote for RFQ: {rfqNumber}",
                            vofh_datum = DateTime.Now.ToString("yyyy-MM-dd"),
                            vofh_geldig_tot = DateTime.Now.AddDays(30).ToString("yyyy-MM-dd"),
                            vofh_status = "OPEN",
                            vofh_prioriteit = firstQuote.Priority ?? "NORMAL",
                            vofh_contact = customerInfo?.CustomerName ?? "Unknown Customer",
                            vofh_memo = "Auto-generated quote from email processing"
                        }
                    }
                        }
                    }
                };

                // Serialize for API call
                var jsonData = JsonSerializer.Serialize(quoteHeaderData, new JsonSerializerOptions { WriteIndented = true });
                LogDebug($"📦 Quote Header data: {jsonData}");

                // 🎯 REAL API CALL: Create quote header
                var content = new StringContent(jsonData, System.Text.Encoding.UTF8, "application/json");
                var responseBody = await _mkgApiClient.PostAsync("Documents/vofh/", content);

                LogDebug($"📥 MKG Quote Header Response: {responseBody}");

                // Extract quote number from response
                var mkgQuoteNumber = ExtractQuoteNumberFromResponse(responseBody);
                var isSuccess = !string.IsNullOrEmpty(mkgQuoteNumber);

                return new MkgQuoteHeaderResult
                {
                    Success = isSuccess,
                    MkgQuoteNumber = mkgQuoteNumber,
                    ErrorMessage = isSuccess ? null : ExtractErrorFromResponse(responseBody),
                    HttpStatusCode = isSuccess ? "200" : "422",
                    RequestPayload = jsonData,
                    ResponsePayload = responseBody
                };
            }
            catch (Exception ex)
            {
                LogDebug($"❌ Error creating MKG Quote Header for {rfqNumber}: {ex.Message}");
                return new MkgQuoteHeaderResult
                {
                    Success = false,
                    ErrorMessage = ex.Message,
                    HttpStatusCode = "EXCEPTION"
                };
            }
        }
        private string ExtractQuoteNumberFromResponse(string responseBody)
        {
            try
            {
                if (string.IsNullOrEmpty(responseBody))
                    return null;

                // Parse JSON response to extract quote number
                using (var document = JsonDocument.Parse(responseBody))
                {
                    var root = document.RootElement;

                    // Check if it's a successful response with quote data
                    if (root.TryGetProperty("response", out var response))
                    {
                        if (response.TryGetProperty("ResultData", out var resultData) && resultData.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var item in resultData.EnumerateArray())
                            {
                                // 🎯 FIX: Look for vofh array in MKG format
                                if (item.TryGetProperty("vofh", out var vofhArray) && vofhArray.ValueKind == JsonValueKind.Array)
                                {
                                    foreach (var vofhItem in vofhArray.EnumerateArray())
                                    {
                                        if (vofhItem.TryGetProperty("vofh_num", out var vofhNum))
                                        {
                                            return vofhNum.GetString();
                                        }
                                    }
                                }

                                // Fallback: direct property
                                if (item.TryGetProperty("vofh_num", out var directVofhNum))
                                {
                                    return directVofhNum.GetString();
                                }
                            }
                        }
                    }

                    // Alternative: look for direct property
                    if (root.TryGetProperty("vofh_num", out var topLevelVofhNum))
                    {
                        return topLevelVofhNum.GetString();
                    }
                }

                LogDebug($"⚠️ Could not extract quote number from response: {responseBody}");
                return null;
            }
            catch (Exception ex)
            {
                LogDebug($"❌ Error extracting quote number: {ex.Message}");
                return null;
            }
        }
        private string ExtractErrorFromResponse(string responseBody)
        {
            try
            {
                if (string.IsNullOrEmpty(responseBody))
                    return "Empty response";

                // Parse JSON response to extract error message
                using (var document = JsonDocument.Parse(responseBody))
                {
                    var root = document.RootElement;

                    // Look for error messages in MKG format
                    if (root.TryGetProperty("response", out var response))
                    {
                        if (response.TryGetProperty("ResultData", out var resultData) && resultData.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var item in resultData.EnumerateArray())
                            {
                                if (item.TryGetProperty("t_messages", out var messages) && messages.ValueKind == JsonValueKind.Array)
                                {
                                    foreach (var message in messages.EnumerateArray())
                                    {
                                        if (message.TryGetProperty("t_melding", out var melding))
                                        {
                                            return melding.GetString();
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Fallback: return the raw response if structured parsing fails
                    return responseBody;
                }
            }
            catch (Exception ex)
            {
                return $"Error parsing response: {ex.Message}";
            }
        }
        private async Task<MkgQuoteLineResult> InjectSingleQuoteLine(QuoteLine quoteLine, string mkgQuoteNumber)
        {
            LogDebug($"🔄 LINE: Injecting quote line {quoteLine.ArtiCode} into quote {mkgQuoteNumber}");

            try
            {
                // TODO: Implement actual quote line injection
                await Task.Delay(100);

                LogDebug($"✅ LINE: Mock quote line injected successfully");

                return new MkgQuoteLineResult
                {
                    Success = true,
                    HttpStatusCode = "200"
                };
            }
            catch (Exception ex)
            {
                LogError($"Quote line injection failed for '{quoteLine.ArtiCode}'", ex);
                return new MkgQuoteLineResult
                {
                    Success = false,
                    ErrorMessage = ex.Message,
                    HttpStatusCode = "500"
                };
            }
        }

        #endregion
    }
}
