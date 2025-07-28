using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using System.Text.Json;
using Mkg_Elcotec_Automation.Models;
using Mkg_Elcotec_Automation.Services;
using System.Net.Http;
using System.Configuration;
using System.Threading;
using System.Text.RegularExpressions;
using System.Globalization;

namespace Mkg_Elcotec_Automation.Controllers
{
    /// <summary>
    /// MKG Order Controller - Complete working version with DebugLogger integration and Live Statistics Support
    /// </summary>
    public class MkgOrderController : MkgController
    {
        #region Configuration Properties

        // Configuration properties from App.config
        private int AdministrationNumber => int.Parse(ConfigurationManager.AppSettings["MkgApi:AdministrationNumber"] ?? "1");
        private int DebtorNumber => int.Parse(ConfigurationManager.AppSettings["MkgApi:DebtorNumber"] ?? "30010");
        private int RelationNumber => int.Parse(ConfigurationManager.AppSettings["MkgApi:RelationNumber"] ?? "2");

        #endregion

        #region Public Methods
        /// <summary>
        /// Inject a single order line into MKG system - FIXED with complete business field mapping
        /// </summary>
        private async Task<MkgOrderResult> InjectSingleOrderLine(
            OrderLine orderLine,
            string mkgOrderNumber,
            EnhancedProgressManager progressManager = null)
        {
            try
            {
                LogDebug($"🔄 Injecting OrderLine {orderLine.ArtiCode} into MKG Order {mkgOrderNumber}");

                // Apply unit conversion
                var originalUnit = orderLine.Unit ?? "PCS";
                var validUnit = ConvertToValidUnit(originalUnit);

                if (originalUnit != validUnit)
                {
                    LogDebug($"🔄 Order unit conversion: '{originalUnit}' → '{validUnit}' for {orderLine.ArtiCode}");
                }

                // ✅ FIXED: Complete business field mapping with all required MKG fields
                var orderLineData = new
                {
                    request = new
                    {
                        InputData = new
                        {
                            vorr = new[]
                            {
                        new
                        {
                            // ✅ CORE IDENTIFICATION FIELDS
                            admi_num = 1,  // Administration number
                            vorh_num = mkgOrderNumber,  // MKG Order number
                            
                            // ✅ ARTICLE & DESCRIPTION - THESE WERE MISSING!
                            vorr_arti_code = orderLine.ArtiCode,  // Article code
                            vorr_oms_1 = string.IsNullOrEmpty(orderLine.Description)
                                ? orderLine.ArtiCode
                                : orderLine.Description,  // Description (was only basic field before)
                            
                            // ✅ QUANTITY & UNIT FIELDS - CRITICAL MISSING DATA!
                            vorr_order_aantal = ParseQuantityToDecimal(orderLine.Quantity),  // Order quantity
                            vorr_eenh_order = validUnit,  // Order unit
                            
                            // ✅ PRICING FIELDS - CRITICAL MISSING DATA!
                            vorr_prijs_order = ParsePriceToDecimal(orderLine.UnitPrice),  // Unit price
                            vorr_totaal_prijs = ParsePriceToDecimal(orderLine.TotalPrice),  // Total price
                            vorr_prijs = ParsePriceToDecimal(orderLine.UnitPrice),  // Price field (alternative)
                            vorr_totaal_excl = ParsePriceToDecimal(orderLine.TotalPrice),  // Total excluding VAT
                            
                            // ✅ DELIVERY & REFERENCE FIELDS - BUSINESS CRITICAL!
                            vorr_gewenste_leverdatum = ParseDateToMkgFormat(orderLine.DeliveryDate),  // Delivery date
                            vorr_leverdatum = ParseDateToMkgFormat(orderLine.DeliveryDate),  // Alternative delivery date field
                            vorr_ref_extern = orderLine.PoNumber,  // External PO reference
                            vorr_memo_extern = orderLine.Notes,  // External notes
                            
                            // ✅ ADDITIONAL BUSINESS FIELDS - IMPORTANT FOR WORKFLOW!
                            vorr_regel = orderLine.LineNumber ?? "001",  // Line number
                            vorr_tekening_nr = orderLine.DrawingNumber,  // Drawing number
                            vorr_revisie = orderLine.Revision ?? "00",  // Revision
                            vorr_leverancier_artikelcode = orderLine.SupplierPartNumber,  // Supplier part number
                            
                            // ✅ WORKFLOW & TRACKING FIELDS
                            vorr_prioriteit = orderLine.Priority ?? "NORMAL",  // Priority
                            vorr_status = orderLine.OrderStatus ?? "OPEN",  // Order status
                            vorr_memo = BuildComprehensiveMemo(orderLine),  // Internal memo with extraction info
                            
                            // ✅ EXTRACTION METADATA - FOR TRACEABILITY
                            vorr_extraction_method = orderLine.ExtractionMethod,  // How this was extracted
                            vorr_bron = "EMAIL_AUTOMATION",  // Source system
                            vorr_verwerkt_door = "ELCOTEC_BOT"  // Processed by
                        }
                    }
                        }
                    }
                };

                // Serialize for API call
                var jsonData = JsonSerializer.Serialize(orderLineData, new JsonSerializerOptions { WriteIndented = true });
                LogDebug($"📦 ENHANCED OrderLine data: {jsonData}");

                // Real API call: Create order line
                var content = new StringContent(jsonData, System.Text.Encoding.UTF8, "application/json");
                var responseBody = await _mkgApiClient.PostAsync("Documents/vorr/", content);

                LogDebug($"📥 MKG Order Line Response: {responseBody}");

                // Check if the response indicates success
                var success = !responseBody.Contains("error") && !responseBody.Contains("\"t_type\":1");

                return new MkgOrderResult
                {
                    ArtiCode = orderLine.ArtiCode,
                    PoNumber = orderLine.PoNumber,
                    Description = orderLine.Description,  // ✅ ADDED: Include description in result
                    Success = success,
                    ErrorMessage = success ? null : ExtractErrorFromResponse(responseBody),
                    MkgOrderId = mkgOrderNumber,
                    ProcessedAt = DateTime.Now,
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
                    progressManager?.IncrementBusinessErrors("MKG Response", errorMessage);
                    LogError($"Business rule violation for {orderLine.ArtiCode}: {errorMessage}", ex);
                }
                else
                {
                    progressManager?.IncrementInjectionErrors();
                    LogError($"Technical error injecting {orderLine.ArtiCode}: {errorMessage}", ex);
                }

                return new MkgOrderResult
                {
                    ArtiCode = orderLine.ArtiCode,
                    PoNumber = orderLine.PoNumber,
                    Description = orderLine.Description,
                    Success = false,
                    ErrorMessage = errorMessage,
                    HttpStatusCode = isBusinessError ? "BUSINESS_RULE_VIOLATION" : "EXCEPTION"
                };
            }
        }

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
                return DateTime.Now.AddDays(14).ToString("yyyy-MM-dd"); // Default to 2 weeks from now

            if (DateTime.TryParse(dateString, out DateTime result))
                return result.ToString("yyyy-MM-dd");

            return DateTime.Now.AddDays(14).ToString("yyyy-MM-dd"); // Fallback
        }

        /// <summary>
        /// Build comprehensive memo with extraction information
        /// </summary>
        private static string BuildComprehensiveMemo(OrderLine orderLine)
        {
            var memo = new List<string>();

            if (!string.IsNullOrEmpty(orderLine.Notes))
                memo.Add($"Notes: {orderLine.Notes}");

            if (!string.IsNullOrEmpty(orderLine.ExtractionMethod))
                memo.Add($"Extracted via: {orderLine.ExtractionMethod}");

            if (!string.IsNullOrEmpty(orderLine.EmailDomain))
                memo.Add($"Source: {orderLine.EmailDomain}");

            memo.Add($"Auto-processed: {DateTime.Now:yyyy-MM-dd HH:mm}");

            return string.Join(" | ", memo);
        }
        /// <summary>
        /// Main injection method with enhanced duplicate handling, logging, and live statistics support
        /// </summary>
        public async Task<MkgOrderInjectionSummary> InjectOrdersAsync(
            List<OrderLine> orderLines,
            IProgress<string> progress = null,
            EnhancedProgressManager progressManager = null)
        {
            var sessionLog = DebugLogger.CreateSessionLog("order_injection");

            var summary = new MkgOrderInjectionSummary
            {
                StartTime = DateTime.Now,
                TotalOrders = 0
            };

            try
            {
                LogDebug($"🔄 Starting MKG order injection for {orderLines.Count} OrderLines");
                progress?.Report($"Starting injection of {orderLines.Count} OrderLines...");

                if (!orderLines.Any())
                {
                    LogDebug("⚠️ No OrderLines provided for injection");
                    return summary;
                }

                // Step 1: Filter valid order lines and group by PO Number
                var validOrderLines = FilterValidOrderLinesWithBusinessValidation(orderLines, progressManager);
                var orderGroups = validOrderLines.GroupBy(ol => ol.PoNumber).ToList();

                LogDebug($"📦 Filtered {orderLines.Count} → {validOrderLines.Count} valid OrderLines");
                LogDebug($"📦 Grouped {validOrderLines.Count} OrderLines into {orderGroups.Count} Order Headers");
                summary.TotalOrders = orderGroups.Count;

                foreach (var orderGroup in orderGroups)
                {
                    var poNumber = orderGroup.Key;
                    var linesInOrder = orderGroup.ToList();

                    try
                    {
                        progress?.Report($"Processing Order: {poNumber} ({linesInOrder.Count} lines)...");
                        LogDebug($"🔄 Processing PO: {poNumber} with {linesInOrder.Count} OrderLines");
                        // Step 2: Check for duplicates
                        var (duplicateExists, existingOrderNumber) = await CheckForExistingOrder(poNumber);
                        if (duplicateExists)
                        {
                            LogDebug($"🔄 DUPLICATE DETECTED: Order for PO {poNumber} already exists as {existingOrderNumber} - skipping creation");
                            progress?.Report($"🔄 Skipping PO {poNumber} - already exists as {existingOrderNumber}");

                            // Mark all lines as skipped duplicates
                            foreach (var line in linesInOrder)
                            {
                                // 🎯 FIX: Add missing duplicate error increment for each duplicate line
                                progressManager?.IncrementDuplicateErrors();

                                summary.OrderResults.Add(new MkgOrderResult
                                {
                                    ArtiCode = line.ArtiCode,
                                    PoNumber = line.PoNumber,
                                    Success = false,
                                    ErrorMessage = $"Order for PO {poNumber} already exists as {existingOrderNumber} - MKG Duplicate",
                                    ProcessedAt = DateTime.Now,
                                    HttpStatusCode = "DUPLICATE_SKIPPED",
                                    MkgOrderId = existingOrderNumber
                                });
                                summary.DuplicatesFiltered++;
                            }

                            summary.SuccessfulInjections++; // Count as successful since order exists
                            continue; // Skip to next order
                        }
                        LogDebug($"✅ No duplicates found for PO: {poNumber} - proceeding with creation");

                        // Step 3: Create Order Header in MKG (only if no duplicate found)
                        progress?.Report($"Creating MKG Order for PO: {poNumber} ({linesInOrder.Count} lines)...");
                        var orderHeaderResult = await CreateMkgOrderHeader(poNumber, linesInOrder.First());

                        if (!orderHeaderResult.Success)
                        {
                            LogDebug($"❌ Failed to create MKG Order Header for PO: {poNumber} - {orderHeaderResult.ErrorMessage}");

                            // If header creation fails, mark all lines as failed
                            foreach (var line in linesInOrder)
                            {
                                summary.OrderResults.Add(new MkgOrderResult
                                {
                                    ArtiCode = line.ArtiCode,
                                    PoNumber = line.PoNumber,
                                    Success = false,
                                    ErrorMessage = $"Order header creation failed: {orderHeaderResult.ErrorMessage}",
                                    ProcessedAt = DateTime.Now,
                                    HttpStatusCode = orderHeaderResult.HttpStatusCode
                                });
                                summary.FailedInjections++;
                                progressManager?.IncrementInjectionErrors();
                            }
                            continue;
                        }

                        // ✅ Order header created successfully
                        var mkgOrderNumber = orderHeaderResult.MkgOrderNumber;
                        LogDebug($"✅ Created NEW MKG Order Header: {mkgOrderNumber} for PO: {poNumber}");

                        progress?.Report($"✅ Created MKG Order: {mkgOrderNumber}, now adding {linesInOrder.Count} lines...");

                        // Step 4: Inject OrderLines
                        foreach (var orderLine in linesInOrder)
                        {
                            try
                            {
                                var lineResult = await InjectSingleOrderLine(orderLine, mkgOrderNumber);
                                summary.OrderResults.Add(lineResult);

                                if (lineResult.Success)
                                {
                                    summary.SuccessfulInjections++;
                                    LogDebug($"✅ OrderLine {orderLine.ArtiCode} → MKG Order {mkgOrderNumber}");
                                }
                                else
                                {
                                    summary.FailedInjections++;
                                    progressManager?.IncrementInjectionErrors();

                                    summary.Errors.Add($"OrderLine {orderLine.ArtiCode}: {lineResult.ErrorMessage}");
                                    LogDebug($"❌ OrderLine {orderLine.ArtiCode} failed: {lineResult.ErrorMessage}");
                                }
                            }
                            catch (Exception lineEx)
                            {
                                summary.FailedInjections++;
                                progressManager?.IncrementInjectionErrors();
                                summary.Errors.Add($"OrderLine {orderLine.ArtiCode}: {lineEx.Message}");
                                LogDebug($"❌ OrderLine {orderLine.ArtiCode} exception: {lineEx.Message}");
                            }
                        }

                    }
                    catch (Exception groupEx)
                    {
                        LogError($"Error processing order group {poNumber}", groupEx);
                        summary.Errors.Add($"Order group {poNumber}: {groupEx.Message}");

                        // Mark all lines in this group as failed
                        foreach (var line in linesInOrder)
                        {
                            summary.OrderResults.Add(new MkgOrderResult
                            {
                                ArtiCode = line.ArtiCode,
                                PoNumber = line.PoNumber,
                                Success = false,
                                ErrorMessage = groupEx.Message,
                                ProcessedAt = DateTime.Now,
                                HttpStatusCode = "GROUP_ERROR"
                            });
                            summary.FailedInjections++;

                            // ✅ INCREMENT ERRORS for live statistics

                            DebugLogger.LogDebug("🔥 DEBUG: IncrementInjectionErrors called!", "MkgOrderController");
                            progressManager?.IncrementInjectionErrors();
                        }
                    }
                }

                summary.EndTime = DateTime.Now;
                summary.ProcessingTime = summary.EndTime - summary.StartTime;

                LogDebug($"🏁 MKG order injection completed: {summary.SuccessfulInjections} successful, {summary.FailedInjections} failed");
                LogDebug($"📊 Duplicate detection statistics: {summary.OrderResults.Count(r => r.HttpStatusCode == "MKG_DUPLICATE_SKIPPED")} duplicates skipped");

                return summary;
            }
            catch (Exception ex)
            {
                LogError("Order injection process failed", ex);
                summary.EndTime = DateTime.Now;
                summary.ProcessingTime = summary.EndTime - summary.StartTime;
                summary.Errors.Add($"Process failed: {ex.Message}");
                return summary;
            }
        }

        /// <summary>
        /// Check for existing orders in MKG to prevent duplicates
        /// </summary>
        private async Task<(bool exists, string existingOrderNumber)> CheckForExistingOrder(string poNumber)
        {
            try
            {
                LogDebug($"🔍 === DUPLICATE CHECK DEBUG START ===");
                LogDebug($"🔍 Checking for existing MKG orders with PO: {poNumber}");

                var cleanPoNumber = poNumber?.Trim();
                if (string.IsNullOrEmpty(cleanPoNumber))
                {
                    LogDebug($"⚠️ Empty PO number provided, skipping duplicate check");
                    return (false, null);
                }

                LogDebug($"🔍 Clean PO Number: '{cleanPoNumber}'");

                // Test different search strategies with detailed logging
                    var searchStrategies = new[]{
                new {
                    Name = "Exact Customer Reference",
                    Filter = $"vorh_ref_uw = \"{cleanPoNumber}\"",
                    Fields = "vorh_num,vorh_ref_uw,vorh_bestelcode_extern,vorh_ref_onze"
                },
                new {
                    Name = "Exact External Order Code",
                    Filter = $"vorh_bestelcode_extern = \"{cleanPoNumber}\"",
                    Fields = "vorh_num,vorh_ref_uw,vorh_bestelcode_extern,vorh_ref_onze"
                },
                new {
                    Name = "Contains Customer Reference",
                    Filter = $"vorh_ref_uw CONTAINS \"{cleanPoNumber}\"",
                    Fields = "vorh_num,vorh_ref_uw,vorh_bestelcode_extern,vorh_ref_onze"
                },
                new {
                    Name = "Contains External Order Code",
                    Filter = $"vorh_bestelcode_extern CONTAINS \"{cleanPoNumber}\"",
                    Fields = "vorh_num,vorh_ref_uw,vorh_bestelcode_extern,vorh_ref_onze"
                }
            };

                for (int i = 0; i < searchStrategies.Length; i++)
                {
                    var strategy = searchStrategies[i];

                    try
                    {
                        LogDebug($"🔍 STRATEGY {i + 1}: {strategy.Name}");
                        LogDebug($"🔍 Filter: {strategy.Filter}");

                        var encodedFilter = Uri.EscapeDataString(strategy.Filter);
                        var endpoint = $"Documents/vorh?Filter={encodedFilter}&FieldList={strategy.Fields}&NumRows=10";

                        LogDebug($"🔗 Full API Endpoint: {endpoint}");

                        var responseBody = await _mkgApiClient.GetAsync(endpoint);

                        LogDebug($"📥 Response Status: {(string.IsNullOrEmpty(responseBody) ? "EMPTY" : "RECEIVED")}");
                        LogDebug($"📥 Response Length: {responseBody?.Length ?? 0} characters");

                        if (!string.IsNullOrEmpty(responseBody))
                        {
                            // Show first 500 characters of response for debugging
                            var preview = responseBody.Length > 500 ? responseBody.Substring(0, 500) + "..." : responseBody;
                            LogDebug($"📥 Response Preview: {preview}");

                            // Check for common error indicators
                            if (responseBody.Contains("\"error\""))
                            {
                                LogDebug($"❌ Response contains error");
                                continue;
                            }

                            if (responseBody.Contains("404") || responseBody.Contains("not found"))
                            {
                                LogDebug($"❌ 404 Not Found response");
                                continue;
                            }

                            // Parse and analyze the response
                            var existingOrders = ParseExistingOrdersResponse(responseBody);
                            LogDebug($"📊 Parsed {existingOrders.Count} existing orders from response");

                            if (existingOrders.Any())
                            {
                                LogDebug($"🔍 Found {existingOrders.Count} existing orders:");
                                foreach (var (orderNumber, poRef) in existingOrders)
                                {
                                    LogDebug($"   📋 Order: {orderNumber} | PO Ref: '{poRef}'");

                                    // Detailed matching logic
                                    bool exactMatch = string.Equals(poRef, cleanPoNumber, StringComparison.OrdinalIgnoreCase);
                                    bool containsMatch = poRef?.Contains(cleanPoNumber, StringComparison.OrdinalIgnoreCase) == true;

                                    LogDebug($"   🔍 Exact Match: {exactMatch} | Contains Match: {containsMatch}");

                                    if (exactMatch || containsMatch)
                                    {
                                        LogDebug($"🎯 DUPLICATE FOUND!");
                                        LogDebug($"🎯 Strategy: {strategy.Name}");
                                        LogDebug($"🎯 Existing Order: {orderNumber}");
                                        LogDebug($"🎯 PO Reference: '{poRef}'");
                                        LogDebug($"🎯 Search PO: '{cleanPoNumber}'");
                                        LogDebug($"🔍 === DUPLICATE CHECK DEBUG END (FOUND) ===");
                                        return (true, orderNumber);
                                    }
                                }
                            }
                            else
                            {
                                LogDebug($"📋 No orders found in this search");
                            }
                        }
                        else
                        {
                            LogDebug($"❌ Empty response from API");
                        }
                    }
                    catch (Exception searchEx)
                    {
                        LogDebug($"❌ Strategy {i + 1} failed: {searchEx.Message}");
                        LogError($"Error in search strategy '{strategy.Name}'", searchEx);
                    }
                }

                LogDebug($"✅ No duplicates found after checking all strategies");
                LogDebug($"🔍 === DUPLICATE CHECK DEBUG END (NOT FOUND) ===");
                return (false, null);
            }
            catch (Exception ex)
            {
                LogDebug($"❌ CRITICAL ERROR in duplicate check: {ex.Message}");
                LogError($"Critical error checking for duplicates for PO {poNumber}", ex);
                LogDebug($"🔍 === DUPLICATE CHECK DEBUG END (ERROR) ===");
                return (false, null); // Assume no duplicates on error to allow injection
            }
        }


        /// <summary>
        /// Parse response from MKG API to extract existing order information
        /// </summary>
        private List<(string orderNumber, string poRef)> ParseExistingOrdersResponse(string responseBody)
        {
            var existingOrders = new List<(string orderNumber, string poRef)>();

            try
            {
                LogDebug($"🔄 === PARSING RESPONSE START ===");

                if (string.IsNullOrEmpty(responseBody))
                {
                    LogDebug($"⚠️ Empty response body for duplicate check");
                    return existingOrders;
                }

                // Handle error responses
                if (responseBody.Contains("\"error\"") || responseBody.Contains("404"))
                {
                    LogDebug($"📋 Error response detected (expected for new POs)");
                    return existingOrders;
                }

                LogDebug($"🔄 Attempting to parse JSON response...");

                using var document = JsonDocument.Parse(responseBody);
                var root = document.RootElement;

                // Log the root structure
                LogDebug($"📋 JSON Root Element Kind: {root.ValueKind}");

                if (root.ValueKind == JsonValueKind.Object)
                {
                    LogDebug($"📋 Root properties found:");
                    foreach (var property in root.EnumerateObject())
                    {
                        LogDebug($"   📋 Property: '{property.Name}' (Type: {property.Value.ValueKind})");
                    }
                }

                // Try multiple parsing strategies
                bool foundData = false;

                // Strategy 1: response.ResultData[].vorh[]
                if (root.TryGetProperty("response", out var response))
                {
                    LogDebug($"📋 Found 'response' property");

                    if (response.TryGetProperty("ResultData", out var resultData) &&
                        resultData.ValueKind == JsonValueKind.Array)
                    {
                        LogDebug($"📋 Found ResultData array with {resultData.GetArrayLength()} items");

                        foreach (var item in resultData.EnumerateArray())
                        {
                            if (item.TryGetProperty("vorh", out var vorhArray) &&
                                vorhArray.ValueKind == JsonValueKind.Array)
                            {
                                LogDebug($"📋 Found vorh array with {vorhArray.GetArrayLength()} orders");
                                foundData = true;

                                foreach (var order in vorhArray.EnumerateArray())
                                {
                                    var vorhNum = GetJsonStringValue(order, "vorh_num");
                                    var poRefUw = GetJsonStringValue(order, "vorh_ref_uw");
                                    var poRefExtern = GetJsonStringValue(order, "vorh_bestelcode_extern");
                                    var poRefOnze = GetJsonStringValue(order, "vorh_ref_onze");

                                    LogDebug($"📋 Raw Order Data:");
                                    LogDebug($"   📋 vorh_num: '{vorhNum}'");
                                    LogDebug($"   📋 vorh_ref_uw: '{poRefUw}'");
                                    LogDebug($"   📋 vorh_bestelcode_extern: '{poRefExtern}'");
                                    LogDebug($"   📋 vorh_ref_onze: '{poRefOnze}'");

                                    if (!string.IsNullOrEmpty(vorhNum))
                                    {
                                        // Use the most relevant PO reference
                                        var poRef = !string.IsNullOrEmpty(poRefUw) ? poRefUw :
                                                   !string.IsNullOrEmpty(poRefExtern) ? poRefExtern :
                                                   poRefOnze;

                                        existingOrders.Add((vorhNum, poRef));
                                        LogDebug($"✅ Added existing order: {vorhNum} with PO ref: '{poRef}'");
                                    }
                                }
                            }
                        }
                    }
                }

                // Strategy 2: Direct data array
                if (!foundData && root.TryGetProperty("data", out var dataArray) &&
                    dataArray.ValueKind == JsonValueKind.Array)
                {
                    LogDebug($"📋 Found direct 'data' array with {dataArray.GetArrayLength()} items");
                    foundData = true;

                    foreach (var order in dataArray.EnumerateArray())
                    {
                        var vorhNum = GetJsonStringValue(order, "vorh_num");
                        var poRefUw = GetJsonStringValue(order, "vorh_ref_uw");
                        var poRefExtern = GetJsonStringValue(order, "vorh_bestelcode_extern");
                        var poRefOnze = GetJsonStringValue(order, "vorh_ref_onze");

                        LogDebug($"📋 Direct Data Order:");
                        LogDebug($"   📋 vorh_num: '{vorhNum}'");
                        LogDebug($"   📋 vorh_ref_uw: '{poRefUw}'");
                        LogDebug($"   📋 vorh_bestelcode_extern: '{poRefExtern}'");
                        LogDebug($"   📋 vorh_ref_onze: '{poRefOnze}'");

                        if (!string.IsNullOrEmpty(vorhNum))
                        {
                            var poRef = !string.IsNullOrEmpty(poRefUw) ? poRefUw :
                                       !string.IsNullOrEmpty(poRefExtern) ? poRefExtern :
                                       poRefOnze;

                            existingOrders.Add((vorhNum, poRef));
                            LogDebug($"✅ Added existing order: {vorhNum} with PO ref: '{poRef}'");
                        }
                    }
                }

                if (!foundData)
                {
                    LogDebug($"❌ No recognizable data structure found");
                    LogDebug($"📋 Response structure doesn't match expected format");

                    // Log first 200 chars for structure analysis
                    var preview = responseBody.Length > 200 ? responseBody.Substring(0, 200) + "..." : responseBody;
                    LogDebug($"📋 Structure analysis preview: {preview}");
                }

                LogDebug($"📊 Total existing orders parsed: {existingOrders.Count}");
                LogDebug($"🔄 === PARSING RESPONSE END ===");

                return existingOrders;
            }
            catch (Exception ex)
            {
                LogDebug($"❌ JSON parsing failed: {ex.Message}");
                LogError($"Failed to parse duplicate check response", ex);
                LogDebug($"📋 Failed response body preview: {responseBody?.Substring(0, Math.Min(responseBody?.Length ?? 0, 300))}");
                LogDebug($"🔄 === PARSING RESPONSE END (ERROR) ===");
                return existingOrders;
            }
        }
        #endregion

        #region Private Helper Methods
        private string GetJsonStringValue(JsonElement element, string propertyName)
        {
            try
            {
                if (element.TryGetProperty(propertyName, out var prop))
                {
                    return prop.ValueKind == JsonValueKind.String ? prop.GetString() : prop.ToString();
                }
                return "";
            }
            catch
            {
                return "";
            }
        }
        /// <summary>
        /// Filter out invalid order lines
        /// </summary>
        private List<OrderLine> FilterValidOrderLines(List<OrderLine> orderLines)
        {
            return orderLines.Where(ol =>
                !string.IsNullOrEmpty(ol.ArtiCode) &&
                !string.IsNullOrEmpty(ol.PoNumber) &&
                ol.ArtiCode != "UNKNOWN-ARTICLE" &&
                !ol.ArtiCode.StartsWith("DELIVER") &&
                !ol.ArtiCode.StartsWith("NUMBER") &&
                !ol.ArtiCode.Contains("UNKNOWN-ARTICLE") &&
                ol.ArtiCode.Length > 2
            ).ToList();
        }

        /// <summary>
        /// Create MKG Order Header for a specific PO
        /// </summary>
        private async Task<MkgOrderHeaderResult> CreateMkgOrderHeader(string poNumber, OrderLine firstOrder)
        {
            try
            {
                LogDebug($"🔄 === BEGIN ORDER HEADER CREATION DEBUG ===");
                LogDebug($"🔄 Creating MKG Order Header for PO: {poNumber}");

                var customerInfo = await FindCustomerByEmailDomain(firstOrder.EmailDomain);

                var orderHeaderData = new
                {
                    request = new
                    {
                        InputData = new
                        {
                            vorh = new[]
                            {
                                new
                                {
                                    admi_num = customerInfo?.AdministrationNumber ?? AdministrationNumber.ToString(),
                                    debi_num = customerInfo?.DebtorNumber ?? DebtorNumber.ToString(),
                                    rela_num = customerInfo?.RelationNumber ?? RelationNumber.ToString(),
                                    vorh_ref_uw = poNumber,
                                    vorh_omschrijving = $"Order for PO: {poNumber}",
                                    vorh_datum = DateTime.Now.ToString("yyyy-MM-dd"),
                                    vorh_gewenste_leverdatum = firstOrder.DeliveryDate ?? DateTime.Now.AddDays(14).ToString("yyyy-MM-dd"),
                                    vorh_status = "OPEN",
                                    vorh_prioriteit = firstOrder.Priority ?? "NORMAL",
                                    vorh_bestelcode_extern = poNumber,
                                    vorh_contact = customerInfo?.CustomerName ?? "Unknown Customer",
                                    vorh_memo = "Auto-generated order from email processing"
                                }
                            }
                        }
                    }
                };

                // Serialize for API call
                var jsonData = JsonSerializer.Serialize(orderHeaderData, new JsonSerializerOptions { WriteIndented = true });
                LogDebug($"📦 Order Header data: {jsonData}");

                // Real API call: Create order header
                var content = new StringContent(jsonData, System.Text.Encoding.UTF8, "application/json");
                var responseBody = await _mkgApiClient.PostAsync("Documents/vorh/", content);

                LogDebug($"📥 MKG Order Header Response: {responseBody}");
                LogDebug($"📥 Response length: {responseBody?.Length ?? 0}");

                // Check if the response indicates success
                var isSuccess = !string.IsNullOrEmpty(responseBody) &&
                               !responseBody.Contains("error") &&
                               !responseBody.Contains("\"t_type\":1") &&
                               !responseBody.Contains("Error");

                LogDebug($"🔍 Success check: {isSuccess}");

                if (isSuccess)
                {
                    var mkgOrderNumber = ExtractOrderNumberFromResponse(responseBody);

                    if (!string.IsNullOrEmpty(mkgOrderNumber))
                    {
                        LogDebug($"✅ Successfully created MKG Order Header: {mkgOrderNumber}");
                        LogDebug($"🔄 === END ORDER HEADER CREATION DEBUG ===");
                        return new MkgOrderHeaderResult
                        {
                            Success = true,
                            MkgOrderNumber = mkgOrderNumber,
                            RequestPayload = jsonData,
                            ResponsePayload = responseBody
                        };
                    }
                    else
                    {
                        LogDebug($"❌ Order header creation succeeded but could not extract order number");
                    }
                }
                else
                {
                    LogDebug($"❌ Order header creation failed - API returned error");
                }

                LogDebug($"🔄 === END ORDER HEADER CREATION DEBUG ===");
                return new MkgOrderHeaderResult
                {
                    Success = false,
                    ErrorMessage = isSuccess ? "Could not extract order number from response" : ExtractErrorFromResponse(responseBody),
                    HttpStatusCode = isSuccess ? "200" : "422",
                    RequestPayload = jsonData,
                    ResponsePayload = responseBody
                };
            }
            catch (Exception ex)
            {
                LogError($"Exception creating MKG order header for {poNumber}", ex);
                return new MkgOrderHeaderResult
                {
                    Success = false,
                    ErrorMessage = ex.Message,
                    HttpStatusCode = "EXCEPTION"
                };
            }
        }
        private bool IsBusinessRuleViolation(string errorMessage)
        {
            if (string.IsNullOrEmpty(errorMessage)) return false;

            var businessRuleKeywords = new[]{
                                            // Duplicate/Validation Errors
                                            "duplicate", "already exists", "invalid reference", "constraint violation",
        
                                            // Business Logic Errors  
                                            "business rule", "validation failed", "unauthorized", "permission denied",
        
                                            // Data Validation Errors
                                            "invalid quantity", "price validation", "invalid unit", "invalid date",
        
                                            // Customer/Authorization Errors
                                            "customer restriction", "not authorized", "access denied", "invalid customer",
        
                                            // MKG Specific Business Rules
                                            "voorraad tekort", "niet toegestaan", "ongeldig", "validatie fout",
                                            "artikel niet gevonden", "klant blokkering", "credit limit"
                                        };

            return businessRuleKeywords.Any(keyword =>
                errorMessage.ToLower().Contains(keyword.ToLower()));
        }

        /// <summary>
        /// Inject a single order line into MKG system
        /// </summary>

        private List<OrderLine> FilterValidOrderLinesWithBusinessValidation(List<OrderLine> orderLines,
                                                                            EnhancedProgressManager progressManager = null)
        {
            var validOrders = new List<OrderLine>();

            foreach (var order in orderLines)
            {
                var businessErrors = ValidateOrderBusinessRules(order);

                if (businessErrors.Any())
                {
                    // This is a business rule violation caught by OUR validation
                    progressManager?.IncrementBusinessErrors("Pre-Validation",
                        $"Order {order.ArtiCode}: {string.Join(", ", businessErrors)}");

                    LogDebug($"🚨 Business rule violation (pre-validation): {order.ArtiCode} - {string.Join(", ", businessErrors)}");
                    continue; // Skip this order
                }

                // Technical validation (not business rules)
                if (string.IsNullOrEmpty(order.ArtiCode) || string.IsNullOrEmpty(order.PoNumber))
                {
                    // This is a technical issue, not business rule
                    continue; // Just skip, don't count as business error
                }

                validOrders.Add(order);
            }

            return validOrders;
        }
        private List<string> ValidateOrderBusinessRules(OrderLine order)
        {
            var errors = new List<string>();

            // Business Rule 1: Quantity must be positive
            if (decimal.TryParse(order.Quantity, out decimal qty) && qty <= 0)
                errors.Add("Quantity must be positive");

            // Business Rule 2: Unit price validation 
            if (decimal.TryParse(order.UnitPrice, out decimal price) && price < 0)
                errors.Add("Unit price cannot be negative");

            // Business Rule 3: Delivery date validation
            if (DateTime.TryParse(order.DeliveryDate, out DateTime deliveryDate) &&
                deliveryDate < DateTime.Now.Date)
                errors.Add("Delivery date cannot be in the past");

            // Business Rule 4: Article code format validation
            if (!string.IsNullOrEmpty(order.ArtiCode) && order.ArtiCode.Length < 3)
                errors.Add("Article code too short");

            // Business Rule 5: Currency/pricing consistency
            if (!string.IsNullOrEmpty(order.UnitPrice) && !string.IsNullOrEmpty(order.TotalPrice))
            {
                if (decimal.TryParse(order.UnitPrice, out decimal unitPrice) &&
                    decimal.TryParse(order.TotalPrice, out decimal totalPrice) &&
                    decimal.TryParse(order.Quantity, out decimal quantity))
                {
                    var calculatedTotal = unitPrice * quantity;
                    if (Math.Abs(calculatedTotal - totalPrice) > 0.01m)
                        errors.Add("Price calculation mismatch");
                }
            }

            return errors;
        }
        /// <summary>
        /// Extract order number from MKG API response
        /// </summary>
        private string ExtractOrderNumberFromResponse(string responseBody)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(responseBody))
                    return null;

                using var document = JsonDocument.Parse(responseBody);
                var root = document.RootElement;

                // Try common structure: response.OutputData.vorh[0].vorh_num
                if (root.TryGetProperty("response", out var response) &&
                    response.TryGetProperty("OutputData", out var outputData) &&
                    outputData.TryGetProperty("vorh", out var vorhArray) &&
                    vorhArray.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in vorhArray.EnumerateArray())
                    {
                        if (item.TryGetProperty("vorh_num", out var vorhNumProp))
                            return vorhNumProp.GetString();
                    }
                }

                // Try alternative: ResultData.vorh[0].vorh_num
                if (root.TryGetProperty("response", out var resp2) &&
                    resp2.TryGetProperty("ResultData", out var resultData) &&
                    resultData.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in resultData.EnumerateArray())
                    {
                        if (item.TryGetProperty("vorh", out var vorhArray2) &&
                            vorhArray2.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var vorhItem in vorhArray2.EnumerateArray())
                            {
                                if (vorhItem.TryGetProperty("vorh_num", out var vorhNum))
                                    return vorhNum.GetString();
                            }
                        }
                    }
                }

                LogDebug($"⚠️ Could not extract order number from response: {responseBody}");
                return null;
            }
            catch (Exception ex)
            {
                LogError($"Error extracting order number", ex);
                return null;
            }
        }

        /// <summary>
        /// Extract error message from MKG API response
        /// </summary>
        private string ExtractErrorFromResponse(string responseBody)
        {
            try
            {
                if (string.IsNullOrEmpty(responseBody))
                    return "Empty response";

                using var document = JsonDocument.Parse(responseBody);
                var root = document.RootElement;

                // Look for error messages in MKG format
                if (root.TryGetProperty("response", out var response))
                {
                    if (response.TryGetProperty("ResultData", out var resultData) &&
                        resultData.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var item in resultData.EnumerateArray())
                        {
                            if (item.TryGetProperty("t_messages", out var messages) &&
                                messages.ValueKind == JsonValueKind.Array)
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
            catch (Exception ex)
            {
                return $"Error parsing response: {ex.Message}";
            }
        }

        #endregion

        #region Logging Methods

        /// <summary>
        /// Debug logging using DebugLogger service
        /// </summary>
        private void LogDebug(string message)
        {
            DebugLogger.LogDebug(message, "MkgOrderController");
        }

        /// <summary>
        /// Error logging using DebugLogger service
        /// </summary>
        private void LogError(string message, Exception ex = null)
        {
            DebugLogger.LogError(message, ex, "MkgOrderController");
        }

        #endregion
    }
}