// ===============================================
// ENHANCED OrderHtmlParser.cs - Better Data Extraction
// Focuses on extracting real data from emails, especially prices
// ===============================================

using System;
using System.Globalization;
using System.IO;
using HtmlAgilityPack;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Linq;
using Mkg_Elcotec_Automation.Models;

namespace Mkg_Elcotec_Automation
{
    class OrderHtmlParser
    {
        public static HtmlDocument Doc { get; set; }
        private static List<string> DebugLog = new List<string>();

        public static void ClearDebugLog() => DebugLog.Clear();
        public static List<string> GetDebugLog() => new List<string>(DebugLog);

        private static void LogDebug(string message)
        {
            var logEntry = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
            DebugLog.Add(logEntry);
            Console.WriteLine($"[ENHANCED_HTML_PARSER] {logEntry}");
        }

        public static HtmlDocument LoadHtml(string body, string folder)
        {
            LogDebug($"Loading HTML content, length: {body?.Length ?? 0}");

            if (string.IsNullOrWhiteSpace(body))
            {
                LogDebug("ERROR: HTML body is null or empty");
                return null;
            }

            try
            {
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(body);
                Doc = doc;

                LogDebug($"HTML document loaded successfully");
                LogDebug($"Document has {doc.DocumentNode?.ChildNodes?.Count ?? 0} child nodes");

                // Enhanced debugging for order extraction
                var orderTable = doc.DocumentNode?.SelectSingleNode("//table[@id='order_lines']");
                LogDebug($"Order lines table found: {orderTable != null}");

                if (orderTable != null)
                {
                    var rows = orderTable.SelectNodes(".//tr");
                    LogDebug($"Order table has {rows?.Count ?? 0} rows");
                }

                var allTables = doc.DocumentNode?.SelectNodes("//table");
                LogDebug($"Total tables in document: {allTables?.Count ?? 0}");

                return doc;
            }
            catch (Exception ex)
            {
                LogDebug($"ERROR loading HTML: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 🔥 ENHANCED: Extract order lines for Weir domain with better price extraction
        /// </summary>
        public List<OrderLine> ExtractOrderLinesWeirDomain()
        {
            LogDebug("=== STARTING ENHANCED WEIR DOMAIN ORDER EXTRACTION ===");

            try
            {
                var orderLines = new List<OrderLine>();

                if (Doc == null)
                {
                    LogDebug("ERROR: Document is null");
                    return orderLines;
                }

                // 🔥 STEP 1: Extract global price data from email text content
                var globalPriceInfo = ExtractGlobalPriceInfo();
                LogDebug($"Global price extracted: Total={globalPriceInfo.totalPrice}, Unit={globalPriceInfo.unitPrice}");

                // 🔥 STEP 2: Extract PO information
                var (poNumber, orderDate) = ExtractPoInfo("");
                LogDebug($"PO Info: Number={poNumber}, Date={orderDate}");

                // 🔥 STEP 3: Find and process order table
                var table = Doc.DocumentNode.SelectSingleNode("//table[@id='order_lines']");
                if (table == null)
                {
                    LogDebug("No table with id 'order_lines' found, trying alternative selectors");

                    // Try alternative table selectors
                    table = Doc.DocumentNode.SelectSingleNode("//table[contains(@class, 'order')]") ??
                           Doc.DocumentNode.SelectSingleNode("//table[.//td[contains(text(), 'Article')]]") ??
                           Doc.DocumentNode.SelectNodes("//table")?.FirstOrDefault();
                }

                if (table == null)
                {
                    LogDebug("ERROR: No suitable table found for order extraction");
                    return orderLines;
                }

                var rows = table.SelectNodes(".//tr");
                LogDebug($"Found {rows?.Count ?? 0} rows in table");

                if (rows == null) return orderLines;

                int rowIndex = 0;
                foreach (var row in rows)
                {
                    rowIndex++;
                    LogDebug($"\n--- Processing Enhanced Row {rowIndex} ---");

                    var columns = row.SelectNodes(".//td");
                    if (columns == null || columns.Count < 3)
                    {
                        LogDebug($"Row {rowIndex}: Insufficient columns ({columns?.Count ?? 0})");
                        continue;
                    }

                    try
                    {
                        // 🔥 ENHANCED: Extract order data using multiple methods
                        var orderData = ExtractEnhancedOrderData(columns, rowIndex, globalPriceInfo);

                        if (string.IsNullOrEmpty(orderData.artiCode) || orderData.artiCode.Contains("SUBCON:"))
                        {
                            LogDebug($"Row {rowIndex}: Invalid article code, skipping");
                            continue;
                        }

                        // 🔥 Create enhanced order line
                        OrderLine order = new OrderLine(
                            FilterInput(columns[0]?.InnerText?.Trim() ?? ""),
                            FilterInput(orderData.artiCode),
                            FilterInput(orderData.description),
                            FilterInput(orderData.drawingNumber),
                            FilterInput(orderData.revision),
                            FilterInput(orderData.supplierPartNumber),
                            "not implemented",
                            FilterInput(orderData.requestedShippingDate),
                            FilterInput(orderData.articleDescription),
                            FormatDate(orderDate),
                            FilterInput(orderData.sapPoLineNumber),
                            FilterInput(orderData.quantity),
                            FilterInput(orderData.unit),
                            orderData.unitPrice,    // 🔥 Don't filter prices - they're already clean
                            orderData.totalPrice    // 🔥 Don't filter prices - they're already clean
                        );

                        orderLines.Add(order);
                        LogDebug($"Row {rowIndex}: ✅ Enhanced order created - Article: {orderData.artiCode}, Unit Price: {orderData.unitPrice}, Total: {orderData.totalPrice}");
                    }
                    catch (Exception ex)
                    {
                        LogDebug($"Row {rowIndex}: ERROR processing: {ex.Message}");
                        continue;
                    }
                }

                LogDebug($"\n=== ENHANCED WEIR EXTRACTION COMPLETE ===");
                LogDebug($"Successfully extracted {orderLines.Count} enhanced order lines");

                return orderLines;
            }
            catch (Exception ex)
            {
                LogDebug($"CRITICAL ERROR in enhanced Weir extraction: {ex.Message}");
                return new List<OrderLine>();
            }
        }

        /// <summary>
        /// 🔥 NEW: Extract global price information from email text content
        /// </summary>
        private (string totalPrice, string unitPrice, int itemCount) ExtractGlobalPriceInfo()
        {
            try
            {
                LogDebug("🔍 Extracting global price info from email text");

                // Get all text content from the document
                var allText = Doc.DocumentNode.InnerText;
                LogDebug($"Email text preview: {allText.Substring(0, Math.Min(300, allText.Length))}...");

                // 🎯 Enhanced Weir-specific price patterns
                var weirPatterns = new[]
                {
                    @"PO\s+Total[:\s]*(\d+)[.,](\d{3})[.,](\d{2})\s*EUR",   // "PO Total: 35.520,00 EUR"
                    @"Total[:\s]*(\d+)[.,](\d{3})[.,](\d{2})\s*EUR",        // "Total: 35.520,00 EUR"  
                    @"(\d+)[.,](\d{3})[.,](\d{2})\s*EUR",                   // "35.520,00 EUR"
                    @"PO\s+Total[:\s]*(\d+)[.,](\d{2})\s*EUR",              // "PO Total: 980,00 EUR"
                    @"Total[:\s]*(\d+)[.,](\d{2})\s*EUR",                   // "Total: 980,00 EUR"
                    @"(\d+)[.,](\d{2})\s*EUR",                              // "980,00 EUR"
                };

                foreach (var pattern in weirPatterns)
                {
                    var match = Regex.Match(allText, pattern, RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        LogDebug($"🎯 PRICE PATTERN MATCHED: {pattern}");

                        string totalPriceStr;
                        if (match.Groups.Count >= 4 && !string.IsNullOrEmpty(match.Groups[3].Value))
                        {
                            // Format with thousands: 35.520,00 → 35520.00
                            totalPriceStr = match.Groups[1].Value + match.Groups[2].Value + "." + match.Groups[3].Value;
                        }
                        else if (match.Groups.Count >= 3 && !string.IsNullOrEmpty(match.Groups[2].Value))
                        {
                            // Simple format: 980,00 → 980.00
                            totalPriceStr = match.Groups[1].Value + "." + match.Groups[2].Value;
                        }
                        else
                        {
                            // Fallback
                            totalPriceStr = match.Groups[1].Value;
                        }

                        LogDebug($"🎯 Reconstructed total price: {totalPriceStr}");

                        if (decimal.TryParse(totalPriceStr, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal totalPrice))
                        {
                            var totalFormatted = totalPrice.ToString("F2", CultureInfo.InvariantCulture);

                            // Try to determine item count and calculate unit price
                            var itemCount = ExtractItemCount(allText);
                            var unitPrice = itemCount > 0 ? (totalPrice / itemCount).ToString("F2", CultureInfo.InvariantCulture) : totalFormatted;

                            LogDebug($"✅ GLOBAL PRICE SUCCESS: Total={totalFormatted}, Unit={unitPrice}, Items={itemCount}");
                            return (totalFormatted, unitPrice, itemCount);
                        }
                    }
                }

                LogDebug("❌ No global price patterns matched");
                return ("0.00", "0.00", 0);
            }
            catch (Exception ex)
            {
                LogDebug($"❌ Error extracting global price info: {ex.Message}");
                return ("0.00", "0.00", 0);
            }
        }

        /// <summary>
        /// 🔥 NEW: Extract item count from email text
        /// </summary>
       private int ExtractItemCount(string text)
        {
            try
            {
                // Look for quantity patterns
                var qtyPatterns = new[]
                {
                    @"(\d+)\s*x\s*\d+\.\d+\.\d+",      // "3 x 891.029.1541"
                    @"Line\s*\d+[:\s]*(\d+)\s*(?:st|pcs|pieces|stuks|each)", // "Line 30: 3 st"
                    @"(?:qty|quantity|aantal)[:\s]*(\d+)", // "Qty: 3"
                    @"(\d+)\s*(?:st|pcs|pieces|stuks|each)", // "3 pcs"
                };

                foreach (var pattern in qtyPatterns)
                {
                    var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                    if (match.Success && int.TryParse(match.Groups[1].Value, out int qty))
                    {
                        LogDebug($"🔢 Found item count: {qty} using pattern: {pattern}");
                        return qty;
                    }
                }

                return 1; // Default to 1 if no quantity found
            }
            catch
            {
                return 1;
            }
        }

        /// <summary>
        /// 🔥 ENHANCED: Extract comprehensive order data from table row
        /// </summary>
        private (string artiCode, string description, string drawingNumber, string revision,
                  string supplierPartNumber, string requestedShippingDate, string articleDescription,
                  string sapPoLineNumber, string quantity, string unit, string unitPrice, string totalPrice)
                 ExtractEnhancedOrderData(HtmlNodeCollection columns, int rowIndex, (string totalPrice, string unitPrice, int itemCount) globalPriceInfo)
        {
            try
            {
                LogDebug($"Row {rowIndex}: Starting enhanced data extraction");

                // 🔥 STEP 1: Extract article code (multiple methods)
                var artiCode = ExtractArticleCode(columns, rowIndex);

                // 🔥 STEP 2: Extract description
                var description = ExtractDescription(columns, rowIndex);

                // 🔥 STEP 3: Extract technical details
                var (drawingNumber, revision) = ExtractTechnicalDetails(columns, rowIndex);

                // 🔥 STEP 4: Extract quantities and units
                var (quantity, unit) = ExtractQuantityAndUnit(columns, rowIndex);

                // 🔥 STEP 5: Extract prices (enhanced with global data) - FIXED LOGIC
                var (unitPrice, totalPrice) = ExtractEnhancedPrices(columns, rowIndex, globalPriceInfo);

                // 🔥 FIX: If we have global price info but extracted prices are 0.00, use global data
                if (unitPrice == "0.00" && totalPrice == "0.00" && globalPriceInfo.totalPrice != "0.00")
                {
                    unitPrice = globalPriceInfo.unitPrice;
                    totalPrice = globalPriceInfo.totalPrice;
                    LogDebug($"Row {rowIndex}: 🎯 APPLIED GLOBAL PRICES - Unit: {unitPrice}, Total: {totalPrice}");
                }

                // 🔥 FIX: Calculate unit price from total if we have quantity
                if (unitPrice == "0.00" && totalPrice != "0.00" && quantity != "1")
                {
                    if (decimal.TryParse(totalPrice, out decimal total) && decimal.TryParse(quantity, out decimal qty) && qty > 0)
                    {
                        var calculatedUnitPrice = (total / qty).ToString("F2", CultureInfo.InvariantCulture);
                        unitPrice = calculatedUnitPrice;
                        LogDebug($"Row {rowIndex}: 🎯 CALCULATED UNIT PRICE - {totalPrice} ÷ {quantity} = {unitPrice}");
                    }
                }

                // 🔥 FIX: Calculate total price from unit if we have quantity
                if (totalPrice == "0.00" && unitPrice != "0.00" && quantity != "1")
                {
                    if (decimal.TryParse(unitPrice, out decimal unit_price) && decimal.TryParse(quantity, out decimal qty) && qty > 0)
                    {
                        var calculatedTotalPrice = (unit_price * qty).ToString("F2", CultureInfo.InvariantCulture);
                        totalPrice = calculatedTotalPrice;
                        LogDebug($"Row {rowIndex}: 🎯 CALCULATED TOTAL PRICE - {unitPrice} × {quantity} = {totalPrice}");
                    }
                }

                LogDebug($"Row {rowIndex}: ✅ Enhanced extraction complete - Article: {artiCode}, Unit Price: {unitPrice}, Total: {totalPrice}");

                return (artiCode, description, drawingNumber, revision, "", "", description,
                       "", quantity, unit, unitPrice, totalPrice);
            }
            catch (Exception ex)
            {
                LogDebug($"Row {rowIndex}: ERROR in enhanced extraction: {ex.Message}");
                return ("", "", "", "", "", "", "", "", "1", "PCS", "0.00", "0.00");
            }
        }

        /// <summary>
        /// 🔥 Extract article code using multiple methods
        /// </summary>
        private string ExtractArticleCode(HtmlNodeCollection columns, int rowIndex)
        {
            try
            {
                // Method 1: From column 2 HTML content (most common)
                if (columns.Count > 2)
                {
                    var dataString = columns[2].InnerHtml;
                    int descriptionIndex = dataString.IndexOf("<");
                    if (descriptionIndex > 0)
                    {
                        string description = dataString.Substring(0, descriptionIndex);
                        string artiCode = description.Split(' ')[0];

                        // Clean up article code
                        if (artiCode.Contains('A'))
                        {
                            artiCode = artiCode.Substring(0, artiCode.IndexOf('A'));
                        }

                        if (!string.IsNullOrEmpty(artiCode) && artiCode.Length >= 3)
                        {
                            LogDebug($"Row {rowIndex}: Article code from method 1: {artiCode}");
                            return artiCode;
                        }
                    }
                }

                // Method 2: From any column containing article-like pattern
                foreach (var column in columns)
                {
                    var text = column.InnerText;
                    var match = Regex.Match(text, @"(\d{3}\.\d{3}\.\d{3,4})", RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        LogDebug($"Row {rowIndex}: Article code from method 2: {match.Groups[1].Value}");
                        return match.Groups[1].Value;
                    }
                }

                LogDebug($"Row {rowIndex}: No article code found");
                return "";
            }
            catch (Exception ex)
            {
                LogDebug($"Row {rowIndex}: Error extracting article code: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// 🔥 Extract description from columns
        /// </summary>
        private string ExtractDescription(HtmlNodeCollection columns, int rowIndex)
        {
            try
            {
                // Try multiple columns for description
                for (int i = 1; i < Math.Min(columns.Count, 4); i++)
                {
                    var text = columns[i].InnerText?.Trim();
                    if (!string.IsNullOrEmpty(text) && text.Length > 10 && text.Length < 200)
                    {
                        LogDebug($"Row {rowIndex}: Description from column {i}: {text.Substring(0, Math.Min(50, text.Length))}...");
                        return text;
                    }
                }

                return "Order item";
            }
            catch
            {
                return "Order item";
            }
        }

        /// <summary>
        /// 🔥 Extract technical details (drawing number, revision)
        /// </summary>
        private (string drawingNumber, string revision) ExtractTechnicalDetails(HtmlNodeCollection columns, int rowIndex)
        {
            try
            {
                string drawingNumber = "";
                string revision = "00";

                // Look for drawing number in HTML content
                foreach (var column in columns)
                {
                    var html = column.InnerHtml;

                    // Drawing number patterns
                    var drawingMatch = Regex.Match(html, @"Drawing\s+(?:Number|No)[:\s]*([A-Z0-9\.\-]+)", RegexOptions.IgnoreCase);
                    if (drawingMatch.Success)
                    {
                        drawingNumber = drawingMatch.Groups[1].Value;
                    }

                    // Revision patterns
                    var revMatch = Regex.Match(html, @"rev[:\s\.]*([A-Z0-9]+)", RegexOptions.IgnoreCase);
                    if (revMatch.Success)
                    {
                        revision = revMatch.Groups[1].Value;
                    }
                }

                LogDebug($"Row {rowIndex}: Technical details - Drawing: {drawingNumber}, Revision: {revision}");
                return (drawingNumber, revision);
            }
            catch
            {
                return ("", "00");
            }
        }

        /// <summary>
        /// 🔥 Extract quantity and unit
        /// </summary>
        private (string quantity, string unit) ExtractQuantityAndUnit(HtmlNodeCollection columns, int rowIndex)
        {
            try
            {
                // Look for quantity in any column
                foreach (var column in columns)
                {
                    var text = column.InnerText;

                    // Quantity patterns
                    var qtyMatch = Regex.Match(text, @"(\d+)\s*(?:x|st|pcs|pieces|stuks|each)", RegexOptions.IgnoreCase);
                    if (qtyMatch.Success)
                    {
                        var qty = qtyMatch.Groups[1].Value;
                        var unitMatch = Regex.Match(text, @"\d+\s*(st|pcs|pieces|stuks|each)", RegexOptions.IgnoreCase);
                        var unit = unitMatch.Success ? unitMatch.Groups[1].Value.ToUpper() : "PCS";

                        LogDebug($"Row {rowIndex}: Quantity: {qty}, Unit: {unit}");
                        return (qty, unit);
                    }
                }

                return ("1", "PCS");
            }
            catch
            {
                return ("1", "PCS");
            }
        }

        /// <summary>
        /// 🔥 ENHANCED: Extract prices using global data and table data
        /// </summary>
        private (string unitPrice, string totalPrice) ExtractEnhancedPrices(HtmlNodeCollection columns, int rowIndex,
             (string totalPrice, string unitPrice, int itemCount) globalPriceInfo)
        {
            try
            {
                LogDebug($"Row {rowIndex}: Starting enhanced price extraction");
                LogDebug($"Row {rowIndex}: Global price info - Total: {globalPriceInfo.totalPrice}, Unit: {globalPriceInfo.unitPrice}, Items: {globalPriceInfo.itemCount}");

                // Method 1: Use global price info if available and significant
                if (globalPriceInfo.totalPrice != "0.00" && decimal.TryParse(globalPriceInfo.totalPrice, out decimal globalTotal) && globalTotal > 0)
                {
                    LogDebug($"Row {rowIndex}: Using global price - Unit: {globalPriceInfo.unitPrice}, Total: {globalPriceInfo.totalPrice}");
                    return (globalPriceInfo.unitPrice, globalPriceInfo.totalPrice);
                }

                // Method 2: Extract from table cells with improved patterns
                string foundUnitPrice = "0.00";
                string foundTotalPrice = "0.00";

                foreach (var column in columns)
                {
                    var text = column.InnerText?.Trim();
                    if (string.IsNullOrEmpty(text)) continue;

                    var price = ExtractPriceFromText(text);
                    if (price != "0.00")
                    {
                        LogDebug($"Row {rowIndex}: Price found in table cell: '{text}' → {price}");

                        // First price found becomes unit price, unless we already have one
                        if (foundUnitPrice == "0.00")
                        {
                            foundUnitPrice = price;
                        }
                        else if (foundTotalPrice == "0.00")
                        {
                            foundTotalPrice = price;
                        }
                    }
                }

                // If we only found one price, use it for both unit and total
                if (foundUnitPrice != "0.00" && foundTotalPrice == "0.00")
                {
                    foundTotalPrice = foundUnitPrice;
                    LogDebug($"Row {rowIndex}: Using single price for both unit and total: {foundUnitPrice}");
                }

                LogDebug($"Row {rowIndex}: Final extracted prices - Unit: {foundUnitPrice}, Total: {foundTotalPrice}");
                return (foundUnitPrice, foundTotalPrice);
            }
            catch (Exception ex)
            {
                LogDebug($"Row {rowIndex}: Error in price extraction: {ex.Message}");
                return ("0.00", "0.00");
            }
        }
        /// <summary>
        /// 🔥 Extract price from text using multiple patterns
        /// </summary>
        private string ExtractPriceFromText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "0.00";

            try
            {
                // Remove extra whitespace and normalize
                text = text.Trim();
                LogDebug($"Extracting price from: '{text}'");

                var pricePatterns = new[]
                {
                    @"€\s*(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})",          // €35.520,00 or €980,00
                    @"(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s*€",          // 35.520,00€ or 980,00€
                    @"€\s*(\d+[.,]\d{2})",                            // €980.00
                    @"(\d+[.,]\d{2})\s*€",                            // 980.00€
                    @"\$\s*(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})",         // $35,520.00
                    @"(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s*\$",         // 35,520.00$
                    @"(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})",              // 35.520,00 or 980,00
                    @"(\d+[.,]\d{2})"                                 // 980.00 or 980,00
                };

                foreach (var pattern in pricePatterns)
                {
                    var match = Regex.Match(text, pattern);
                    if (match.Success)
                    {
                        var priceStr = match.Groups[1].Value;
                        var cleanPrice = CleanAndValidatePrice(priceStr);
                        if (cleanPrice != "0.00")
                        {
                            LogDebug($"✅ Price pattern matched: '{pattern}' → '{priceStr}' → {cleanPrice}");
                            return cleanPrice;
                        }
                    }
                }

                LogDebug($"❌ No price patterns matched in: '{text}'");
            }
            catch (Exception ex)
            {
                LogDebug($"❌ Error extracting price from '{text}': {ex.Message}");
            }

            return "0.00";
        }
        /// <summary>
        /// 🔥 Clean and validate price values
        /// </summary>
        private string CleanAndValidatePrice(string price)
        {
            if (string.IsNullOrWhiteSpace(price)) return "0.00";

            try
            {
                price = price.Replace("€", "").Replace("$", "").Replace("£", "").Trim();

                // Handle European format (comma as decimal separator)
                if (price.Contains(",") && !price.Contains("."))
                {
                    price = price.Replace(",", ".");
                }
                else if (price.Contains(",") && price.Contains("."))
                {
                    // Handle thousands separator: 1.234,56 -> 1234.56
                    var lastComma = price.LastIndexOf(',');
                    var lastDot = price.LastIndexOf('.');
                    if (lastComma > lastDot)
                    {
                        price = price.Replace(".", "").Replace(",", ".");
                    }
                }

                // Remove any remaining non-numeric characters except decimal point
                price = Regex.Replace(price, @"[^\d\.]", "");

                if (decimal.TryParse(price, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result) && result > 0)
                {
                    return result.ToString("F2", CultureInfo.InvariantCulture);
                }
            }
            catch
            {
                // Fall through to return 0.00
            }

            return "0.00";
        }

        // ===============================================
        // EXISTING HELPER METHODS (Enhanced)
        // ===============================================

        public static string FormatDate(string dateStr, string currentFormat = "dd/MM/yyyy", string desiredFormat = "dd-MM-yyyy")
        {
            if (string.IsNullOrWhiteSpace(dateStr))
            {
                return DateTime.Now.ToString(desiredFormat);
            }

            try
            {
                DateTime date = DateTime.ParseExact(dateStr.Trim(), currentFormat, CultureInfo.InvariantCulture);
                return date.ToString(desiredFormat);
            }
            catch (FormatException)
            {
                return DateTime.Now.ToString(desiredFormat);
            }
        }

        public (string poNumber, string OrderDate) ExtractPoInfo(string body)
        {
            LogDebug("🔍 Extracting PO info from document");

            try
            {
                var allText = Doc.DocumentNode.InnerText;

                // Enhanced PO number patterns for Weir
                var poPatterns = new[]
                {
                    @"Purchase\s+Order\s+#(\d{10})",                    // "Purchase Order #4501533672"
                    @"Purchase\s+Order\s+(\d{10})",
                    @"PO[\s#:]*(\d{10})",
                    @"PO[\s#:]*(\d{8,12})",
                };

                foreach (var pattern in poPatterns)
                {
                    var match = Regex.Match(allText, pattern, RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        var poNumber = "PO-" + match.Groups[1].Value;
                        LogDebug($"✅ PO Number extracted: {poNumber}");

                        // Try to extract date
                        var dateMatch = Regex.Match(allText, @"(\d{2}/\d{2}/\d{4})");
                        var orderDate = dateMatch.Success ? dateMatch.Groups[1].Value : DateTime.Now.ToString("dd/MM/yyyy");

                        return (poNumber, orderDate);
                    }
                }

                LogDebug("❌ No PO number found");
                return ("PO-AUTO-" + DateTime.Now.ToString("yyyyMMddHHmm"), DateTime.Now.ToString("dd/MM/yyyy"));
            }
            catch (Exception ex)
            {
                LogDebug($"ERROR extracting PO info: {ex.Message}");
                return ("PO-ERROR", DateTime.Now.ToString("dd/MM/yyyy"));
            }
        }

        /// <summary>
        /// Filter and clean input text
        /// </summary>
        private static string FilterInput(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "";

            return input.Trim()
                       .Replace("\r\n", " ")
                       .Replace("\n", " ")
                       .Replace("\t", " ")
                       .Replace("&nbsp;", " ")
                       .Trim();
        }

        // ===============================================
        // BACKWARD COMPATIBILITY METHODS
        // ===============================================

        public List<OrderLine> ExtractOrderLinesOutlookForwardMail()
        {
            LogDebug("=== FALLBACK: Using enhanced extraction for Outlook forward mail ===");
            return ExtractOrderLinesWeirDomain(); // Use enhanced method as fallback
        }
    }
}