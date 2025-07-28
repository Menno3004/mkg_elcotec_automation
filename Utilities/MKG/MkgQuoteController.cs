using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using Mkg_Elcotec_Automation.Services;
using System.Configuration;
using Mkg_Elcotec_Automation.Models;
using System.Text.Json;
using System.Net.Http;
using System.Globalization;

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
        /// <summary>
        /// Inject a single quote line into MKG system - FIXED with complete business field mapping
        /// </summary>
        /// <summary>
        /// Inject a single quote line into MKG system - COMPLETE REPLACEMENT
        /// </summary>
        private async Task<MkgQuoteLineResult> InjectSingleQuoteLine(
            QuoteLine quoteLine,
            string mkgQuoteNumber,
            EnhancedProgressManager progressManager = null)
        {
            LogDebug($"🔄 LINE: Injecting quote line {quoteLine.ArtiCode} into quote {mkgQuoteNumber}");

            try
            {
                // Apply unit conversion
                var originalUnit = quoteLine.Unit ?? "PCS";
                var validUnit = ConvertToValidUnit(originalUnit);

                if (originalUnit != validUnit)
                {
                    LogDebug($"🔄 Quote unit conversion: '{originalUnit}' → '{validUnit}' for {quoteLine.ArtiCode}");
                }

                // ✅ FIXED: Complete business field mapping with all required MKG quote fields
                var quoteLineData = new
                {
                    request = new
                    {
                        InputData = new
                        {
                            vofr = new[]  // ✅ Note: vofr for quotes, vorr for orders
                            {
                        new
                        {
                            // ✅ CORE IDENTIFICATION FIELDS
                            admi_num = 1,  // Administration number
                            vofh_num = mkgQuoteNumber,  // MKG Quote number
                            
                            // ✅ ARTICLE & DESCRIPTION - THESE WERE MISSING!
                            vofr_arti_code = quoteLine.ArtiCode,  // Article code
                            vofr_oms_1 = string.IsNullOrEmpty(quoteLine.Description)
                                ? quoteLine.ArtiCode
                                : quoteLine.Description,  // Description
                            
                            // ✅ QUANTITY & UNIT FIELDS - CRITICAL MISSING DATA!
                            vofr_order_aantal = ParseQuantityToDecimal(quoteLine.Quantity),  // Quote quantity
                            vofr_eenh_order = validUnit,  // Quote unit
                            
                            // ✅ PRICING FIELDS - CRITICAL MISSING DATA FOR QUOTES!
                            vofr_prijs_order = ParsePriceToDecimal(quoteLine.QuotedPrice),  // Quoted price per unit
                            vofr_totaal_prijs = quoteLine.TotalPrice > 0 ? quoteLine.TotalPrice : ParsePriceToDecimal(quoteLine.QuotedPrice),  // Total quoted price
                            vofr_prijs = ParsePriceToDecimal(quoteLine.QuotedPrice),  // Price field (alternative)
                            vofr_totaal_excl = quoteLine.TotalPrice > 0 ? quoteLine.TotalPrice : ParsePriceToDecimal(quoteLine.QuotedPrice),  // Total excluding VAT
                            
                            // ✅ QUOTE-SPECIFIC FIELDS - BUSINESS CRITICAL FOR QUOTES!
                            vofr_levertijd = SafeGetStringValue(quoteLine.LeadTime, "14"),  // Lead time in days
                            vofr_geldig_tot = ParseDateToMkgFormat(quoteLine.ValidUntil),  // Quote validity date
                            vofr_leverwijze = "STANDARD",  // Delivery method (default)
                            
                            // ✅ DELIVERY & REFERENCE FIELDS
                            vofr_gewenste_leverdatum = ParseDateToMkgFormat(quoteLine.RequestedDeliveryDate),  // Requested delivery
                            vofr_ref_extern = quoteLine.RfqNumber,  // External RFQ reference
                            vofr_memo_extern = SafeGetStringValue(quoteLine.Notes, ""),  // External notes/requirements
                            
                            // ✅ ADDITIONAL BUSINESS FIELDS - IMPORTANT FOR WORKFLOW!
                            vofr_regel = quoteLine.LineNumber ?? "001",  // Line number
                            vofr_tekening_nr = quoteLine.DrawingNumber,  // Drawing number
                            vofr_revisie = quoteLine.Revision ?? "00",  // Revision
                            vofr_leverancier_artikelcode = quoteLine.CustomerPartNumber,  // Customer part number (closest match)
                            
                            // ✅ QUOTE WORKFLOW & TRACKING FIELDS
                            vofr_prioriteit = quoteLine.Priority ?? "NORMAL",  // Priority level
                            vofr_status = quoteLine.QuoteStatus ?? "PENDING",  // Quote status
                            vofr_memo = BuildComprehensiveQuoteMemo(quoteLine),  // Internal memo with extraction info
                            
                            // ✅ QUOTE-SPECIFIC BUSINESS FIELDS (using available fields)
                            vofr_min_aantal = ParseQuantityToDecimal(quoteLine.Quantity),  // Use same quantity for min
                            
                            // ✅ EXTRACTION METADATA - FOR TRACEABILITY
                            vofr_extraction_method = quoteLine.ExtractionMethod,  // How this was extracted
                            vofr_bron = "EMAIL_AUTOMATION",  // Source system
                            vofr_verwerkt_door = "ELCOTEC_BOT",  // Processed by
                            
                            // ✅ QUOTE VALIDITY & TERMS (using available fields)
                            vofr_voorwaarden = SafeGetStringValue(quoteLine.PaymentTerms, "Standard terms apply"),  // Payment terms
                            vofr_opmerkingen = SafeGetStringValue(quoteLine.SpecialInstructions, "")  // Special instructions
                        }
                    }
                        }
                    }
                };

                // Serialize for API call
                var jsonData = JsonSerializer.Serialize(quoteLineData, new JsonSerializerOptions { WriteIndented = true });
                LogDebug($"📦 ENHANCED Quote Line data: {jsonData}");

                // ✅ REAL API call: Create quote line (was previously just a mock)
                var content = new StringContent(jsonData, System.Text.Encoding.UTF8, "application/json");
                var responseBody = await _mkgApiClient.PostAsync("Documents/vofr/", content);

                LogDebug($"📥 MKG Quote Line Response: {responseBody}");

                // Check if the response indicates success
                var success = !responseBody.Contains("error") && !responseBody.Contains("\"t_type\":1");

                if (success)
                {
                    LogDebug($"✅ LINE: Quote line {quoteLine.ArtiCode} injected successfully into quote {mkgQuoteNumber}");
                }
                else
                {
                    LogDebug($"❌ LINE: Quote line {quoteLine.ArtiCode} injection failed: {ExtractErrorFromResponse(responseBody)}");
                }

                return new MkgQuoteLineResult
                {
                    Success = success,
                    ErrorMessage = success ? null : ExtractErrorFromResponse(responseBody),
                    HttpStatusCode = success ? "200" : "422",
                    RequestPayload = jsonData,
                    ResponsePayload = responseBody
                };
            }
            catch (Exception ex)
            {
                var errorMessage = ex.Message;
                var isBusinessError = IsBusinessRuleViolation(errorMessage);

                if (isBusinessError)
                {
                    progressManager?.IncrementBusinessErrors("MKG Quote Response", errorMessage);
                    LogError($"Business rule violation for quote {quoteLine.ArtiCode}: {errorMessage}", ex);
                }
                else
                {
                    progressManager?.IncrementInjectionErrors();
                    LogError($"Technical error injecting quote {quoteLine.ArtiCode}: {errorMessage}", ex);
                }

                return new MkgQuoteLineResult
                {
                    Success = false,
                    ErrorMessage = errorMessage,
                    HttpStatusCode = isBusinessError ? "BUSINESS_RULE_VIOLATION" : "EXCEPTION",
                    RequestPayload = "Error occurred before creating payload",
                    ResponsePayload = null
                };
            }
        }

        #region Helper Methods for Quote Processing

        /// <summary>
        /// Convert quantity string to decimal for MKG API
        /// </summary>
        private static decimal ParseQuantityToDecimal(string quantity)
        {
            if (string.IsNullOrEmpty(quantity)) return 1.0m;

            // Clean the quantity string
            var cleanQty = quantity.Replace(",", ".").Trim();

            if (decimal.TryParse(cleanQty, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
                return result;

            return 1.0m; // Default quantity
        }

        /// <summary>
        /// Convert price string to decimal for MKG API
        /// </summary>
        private static decimal? ParsePriceToDecimal(string price)
        {
            if (string.IsNullOrEmpty(price)) return null;

            // Clean the price string (remove currency symbols, etc.)
            var cleanPrice = price.Replace("€", "").Replace("$", "").Replace(",", ".").Trim();

            if (decimal.TryParse(cleanPrice, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
                return result;

            return null; // No price available
        }

        /// <summary>
        /// Convert date string to MKG format (yyyy-MM-dd)
        /// </summary>
        private static string ParseDateToMkgFormat(string dateString)
        {
            if (string.IsNullOrEmpty(dateString))
                return DateTime.Now.AddDays(30).ToString("yyyy-MM-dd"); // Default to 30 days from now for quotes

            if (DateTime.TryParse(dateString, out DateTime result))
                return result.ToString("yyyy-MM-dd");

            return DateTime.Now.AddDays(30).ToString("yyyy-MM-dd"); // Fallback for quotes
        }

        /// <summary>
        /// Safely get string value from dynamic property or return default
        /// </summary>
        private string SafeGetStringValue(dynamic value, string defaultValue)
        {
            try
            {
                return value?.ToString() ?? defaultValue;
            }
            catch
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// Build comprehensive memo with extraction information for quotes
        /// </summary>
        private string BuildComprehensiveQuoteMemo(QuoteLine quoteLine)
        {
            var memo = new List<string>();

            if (!string.IsNullOrEmpty(quoteLine.QuoteNotes))
                memo.Add($"Notes: {quoteLine.QuoteNotes}");

            if (!string.IsNullOrEmpty(quoteLine.ExtractionMethod))
                memo.Add($"Extracted via: {quoteLine.ExtractionMethod}");

            if (!string.IsNullOrEmpty(quoteLine.EmailDomain))
                memo.Add($"Source: {quoteLine.EmailDomain}");

            if (!string.IsNullOrEmpty(quoteLine.SpecialInstructions))
                memo.Add($"Special: {quoteLine.SpecialInstructions}");

            memo.Add($"Quote auto-processed: {DateTime.Now:yyyy-MM-dd HH:mm}");

            return string.Join(" | ", memo);
        }

        /// <summary>
        /// Check if error message indicates a business rule violation
        /// </summary>
        private bool IsBusinessRuleViolation(string errorMessage)
        {
            if (string.IsNullOrEmpty(errorMessage)) return false;

            var businessRuleKeywords = new[]
            {
        // Duplicate/Validation Errors
        "duplicate", "already exists", "invalid reference", "constraint violation",

        // Business Logic Errors  
        "business rule", "validation failed", "unauthorized", "permission denied",

        // Data Validation Errors
        "invalid quantity", "price validation", "invalid unit", "invalid date",

        // Customer/Authorization Errors
        "customer restriction", "not authorized", "access denied", "invalid customer",

        // MKG Specific Business Rules for Quotes
        "quote expired", "invalid quote", "quote already processed", "price not allowed",
        "voorraad tekort", "niet toegestaan", "ongeldig", "validatie fout",
        "artikel niet gevonden", "klant blokkering", "credit limit",
        "offerte verlopen", "offerte ongeldig"
    };

            return businessRuleKeywords.Any(keyword =>
                errorMessage.ToLower().Contains(keyword.ToLower()));
        }

        #endregion

        /// <summary>
        /// Parse discount percentage for MKG API
        /// </summary>
        private static decimal? ParseDiscountPercentage(string discountPercentage)
        {
            if (string.IsNullOrEmpty(discountPercentage)) return null;

            var cleanDiscount = discountPercentage.Replace("%", "").Replace(",", ".").Trim();

            if (decimal.TryParse(cleanDiscount, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
                return result;

            return null;
        }

        /// <summary>
        /// Parse surcharge percentage for MKG API
        /// </summary>
        private static decimal? ParseSurchargePercentage(string surchargePercentage)
        {
            if (string.IsNullOrEmpty(surchargePercentage)) return null;

            var cleanSurcharge = surchargePercentage.Replace("%", "").Replace(",", ".").Trim();

            if (decimal.TryParse(cleanSurcharge, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
                return result;

            return null;
        }
   
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
        #endregion
    }
}
