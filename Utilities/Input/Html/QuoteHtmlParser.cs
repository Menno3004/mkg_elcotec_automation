// ===============================================
// ENHANCED QuoteHtmlParser.cs - Better Quote Data Extraction
// Focuses on extracting real quote data from emails, especially prices
// ===============================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Mkg_Elcotec_Automation.Models;

namespace Mkg_Elcotec_Automation.Utilities.Input.Html
{
    public class QuoteHtmlParser
    {
        public static HtmlDocument Doc { get; set; }
        private static List<string> DebugLog = new List<string>();

        public static void ClearDebugLog() => DebugLog.Clear();
        public static List<string> GetDebugLog() => new List<string>(DebugLog);

        private static void LogDebug(string message)
        {
            var logEntry = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
            DebugLog.Add(logEntry);
            Console.WriteLine($"[ENHANCED_QUOTE_PARSER] {logEntry}");
        }

        /// <summary>
        /// 🔥 ENHANCED: Extract quote lines for Weir domain with better price extraction
        /// </summary>
        public List<QuoteLine> ExtractQuoteLinesWeirDomain()
        {
            LogDebug("=== STARTING ENHANCED WEIR DOMAIN QUOTE EXTRACTION ===");

            try
            {
                var quoteLines = new List<QuoteLine>();

                if (Doc == null)
                {
                    LogDebug("ERROR: Document is null");
                    return quoteLines;
                }

                // 🔥 STEP 1: Extract global price and RFQ info from email text
                var globalQuoteInfo = ExtractGlobalQuoteInfo();
                LogDebug($"Global quote info: RFQ={globalQuoteInfo.rfqNumber}, Price={globalQuoteInfo.totalPrice}");

                // 🔥 STEP 2: Find and process quote/order table (quotes often use order tables)
                var table = Doc.DocumentNode.SelectSingleNode("//table[@id='quote_lines']") ??
                           Doc.DocumentNode.SelectSingleNode("//table[@id='rfq_lines']") ??
                           Doc.DocumentNode.SelectSingleNode("//table[@id='order_lines']") ??
                           Doc.DocumentNode.SelectSingleNode("//table[contains(@class, 'quote')]") ??
                           Doc.DocumentNode.SelectSingleNode("//table[.//td[contains(text(), 'RFQ')]]") ??
                           Doc.DocumentNode.SelectNodes("//table")?.FirstOrDefault();

                if (table == null)
                {
                    LogDebug("ERROR: No suitable table found for quote extraction");
                    return quoteLines;
                }

                var rows = table.SelectNodes(".//tr");
                LogDebug($"Found {rows?.Count ?? 0} rows in quote table");

                if (rows == null) return quoteLines;

                int rowIndex = 0;
                foreach (var row in rows)
                {
                    rowIndex++;
                    LogDebug($"\n--- Processing Enhanced Quote Row {rowIndex} ---");

                    var columns = row.SelectNodes(".//td");
                    if (columns == null || columns.Count < 3)
                    {
                        LogDebug($"Row {rowIndex}: Insufficient columns ({columns?.Count ?? 0})");
                        continue;
                    }

                    try
                    {
                        // 🔥 ENHANCED: Extract quote data using multiple methods
                        var quoteData = ExtractEnhancedQuoteData(columns, rowIndex, globalQuoteInfo);

                        if (string.IsNullOrEmpty(quoteData.artiCode) || quoteData.artiCode.Contains("SUBCON:"))
                        {
                            LogDebug($"Row {rowIndex}: Invalid article code, skipping");
                            continue;
                        }

                        // 🔥 Create enhanced quote line
                        var quote = new QuoteLine(
                            artiCode: quoteData.artiCode,
                            description: quoteData.description,
                            rfqNumber: quoteData.rfqNumber,
                            drawingNumber: quoteData.drawingNumber,
                            revision: quoteData.revision,
                            requestedDeliveryDate: quoteData.requestedDeliveryDate,
                            quoteDate: quoteData.quoteDate,
                            quotedPrice: quoteData.quotedPrice,
                            validUntil: quoteData.validUntil,
                            quoteStatus: "Draft",
                            priority: DeterminePriority(quoteData.quotedPrice),
                            extractionMethod: "ENHANCED_WEIR_HTML",
                            extractionDomain: "weir.com"
                        );

                        quoteLines.Add(quote);
                        LogDebug($"Row {rowIndex}: ✅ Enhanced quote created - Article: {quoteData.artiCode}, Price: {quoteData.quotedPrice}");
                    }
                    catch (Exception ex)
                    {
                        LogDebug($"Row {rowIndex}: ERROR processing: {ex.Message}");
                        continue;
                    }
                }

                LogDebug($"\n=== ENHANCED WEIR QUOTE EXTRACTION COMPLETE ===");
                LogDebug($"Successfully extracted {quoteLines.Count} enhanced quote lines");

                return quoteLines;
            }
            catch (Exception ex)
            {
                LogDebug($"CRITICAL ERROR in enhanced Weir quote extraction: {ex.Message}");
                return new List<QuoteLine>();
            }
        }

        /// <summary>
        /// 🔥 NEW: Extract global quote information from email text content
        /// </summary>
        private (string rfqNumber, string totalPrice, string quoteDate, string validUntil) ExtractGlobalQuoteInfo()
        {
            try
            {
                LogDebug("🔍 Extracting global quote info from email text");

                var allText = Doc.DocumentNode.InnerText;
                LogDebug($"Email text preview: {allText.Substring(0, Math.Min(300, allText.Length))}...");

                // Extract RFQ number
                var rfqNumber = ExtractRfqNumber(allText);

                // Extract total price (same patterns as orders)
                var totalPrice = ExtractTotalPrice(allText);

                // Extract dates
                var quoteDate = DateTime.Now.ToString("dd-MM-yyyy");
                var validUntil = DateTime.Now.AddDays(30).ToString("dd-MM-yyyy");

                LogDebug($"✅ Global quote info - RFQ: {rfqNumber}, Price: {totalPrice}");
                return (rfqNumber, totalPrice, quoteDate, validUntil);
            }
            catch (Exception ex)
            {
                LogDebug($"❌ Error extracting global quote info: {ex.Message}");
                return ("", "0.00", DateTime.Now.ToString("dd-MM-yyyy"), DateTime.Now.AddDays(30).ToString("dd-MM-yyyy"));
            }
        }

        /// <summary>
        /// 🔥 Extract RFQ number from text
        /// </summary>
        private string ExtractRfqNumber(string text)
        {
            var rfqPatterns = new[]
            {
                @"RFQ[:\s#]*([A-Z0-9\-]{6,20})",
                @"Request\s+for\s+Quote[:\s#]*([A-Z0-9\-]{6,20})",
                @"Quote\s+Request[:\s#]*([A-Z0-9\-]{6,20})",
                @"Purchase\s+Order\s+#(\d{10})", // Sometimes RFQ uses PO format
            };

            foreach (var pattern in rfqPatterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var rfqNumber = "RFQ-" + match.Groups[1].Value;
                    LogDebug($"🎯 RFQ Number found: {rfqNumber}");
                    return rfqNumber;
                }
            }

            // Generate RFQ number based on current time
            var generatedRfq = $"RFQ-AUTO-{DateTime.Now:yyyyMMddHHmm}";
            LogDebug($"🎯 Generated RFQ Number: {generatedRfq}");
            return generatedRfq;
        }

        /// <summary>
        /// 🔥 Extract total price using same patterns as orders
        /// </summary>
        private string ExtractTotalPrice(string text)
        {
            var pricePatterns = new[]
            {
                @"(?:Total|Quote|Price)[:\s]*(\d+[.,]\d{3})[.,](\d{2})\s*EUR",   // "Total: 35.520,00 EUR"
                @"(?:Total|Quote|Price)[:\s]*(\d+)[.,](\d{2})\s*EUR",            // "Total: 980,00 EUR"
                @"(\d+[.,]\d{3})[.,](\d{2})\s*EUR",                              // "35.520,00 EUR"
                @"(\d+)[.,](\d{2})\s*EUR",                                       // "980,00 EUR"
            };

            foreach (var pattern in pricePatterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    string totalPriceStr;
                    if (match.Groups.Count >= 3 && !string.IsNullOrEmpty(match.Groups[2].Value))
                    {
                        // Format with thousands: 35.520,00 → 35520.00
                        totalPriceStr = match.Groups[1].Value.Replace(",", "").Replace(".", "") + "." + match.Groups[2].Value;
                    }
                    else
                    {
                        // Simple format: 980,00 → 980.00
                        totalPriceStr = match.Groups[1].Value.Replace(",", ".");
                    }

                    if (decimal.TryParse(totalPriceStr, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal totalPrice))
                    {
                        var formatted = totalPrice.ToString("F2", CultureInfo.InvariantCulture);
                        LogDebug($"🎯 Total price found: {formatted}");
                        return formatted;
                    }
                }
            }

            return "0.00";
        }

        /// <summary>
        /// 🔥 ENHANCED: Extract comprehensive quote data from table row
        /// </summary>
        private (string artiCode, string description, string rfqNumber, string drawingNumber, string revision,
                 string requestedDeliveryDate, string quoteDate, string quotedPrice, string totalPrice, string leadTime, string validUntil)
                ExtractEnhancedQuoteData(HtmlNodeCollection columns, int rowIndex, (string rfqNumber, string totalPrice, string quoteDate, string validUntil) globalQuoteInfo)
        {
            try
            {
                LogDebug($"Row {rowIndex}: Starting enhanced quote data extraction");

                // 🔥 Extract article code (same logic as orders)
                var artiCode = ExtractArticleCode(columns, rowIndex);

                // 🔥 Extract description
                var description = ExtractDescription(columns, rowIndex);

                // 🔥 Extract technical details
                var (drawingNumber, revision) = ExtractTechnicalDetails(columns, rowIndex);

                // 🔥 Extract prices (enhanced with global data)
                var (quotedPrice, totalPrice) = ExtractEnhancedQuotePrices(columns, rowIndex, globalQuoteInfo);

                // 🔥 Extract dates and lead time
                var leadTime = ExtractLeadTime(columns, rowIndex);
                var requestedDeliveryDate = DateTime.Now.AddDays(14).ToString("dd-MM-yyyy");

                LogDebug($"Row {rowIndex}: ✅ Enhanced quote extraction complete - Article: {artiCode}, Price: {quotedPrice}");

                return (artiCode, description, globalQuoteInfo.rfqNumber, drawingNumber, revision,
                       requestedDeliveryDate, globalQuoteInfo.quoteDate, quotedPrice, totalPrice, leadTime, globalQuoteInfo.validUntil);
            }
            catch (Exception ex)
            {
                LogDebug($"Row {rowIndex}: ERROR in enhanced quote extraction: {ex.Message}");
                return ("", "", "", "", "", "", "", "0.00", "0.00", "14", "");
            }
        }

        /// <summary>
        /// 🔥 Extract article code (reuse logic from order parser)
        /// </summary>
        private string ExtractArticleCode(HtmlNodeCollection columns, int rowIndex)
        {
            try
            {
                // Method 1: From column 2 HTML content
                if (columns.Count > 2)
                {
                    var dataString = columns[2].InnerHtml;
                    int descriptionIndex = dataString.IndexOf("<");
                    if (descriptionIndex > 0)
                    {
                        string description = dataString.Substring(0, descriptionIndex);
                        string artiCode = description.Split(' ')[0];

                        if (artiCode.Contains('A'))
                        {
                            artiCode = artiCode.Substring(0, artiCode.IndexOf('A'));
                        }

                        if (!string.IsNullOrEmpty(artiCode) && artiCode.Length >= 3)
                        {
                            LogDebug($"Row {rowIndex}: Article code: {artiCode}");
                            return artiCode;
                        }
                    }
                }

                // Method 2: Pattern matching in any column
                foreach (var column in columns)
                {
                    var text = column.InnerText;
                    var match = Regex.Match(text, @"(\d{3}\.\d{3}\.\d{3,4})", RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        LogDebug($"Row {rowIndex}: Article code from pattern: {match.Groups[1].Value}");
                        return match.Groups[1].Value;
                    }
                }

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
                for (int i = 1; i < Math.Min(columns.Count, 4); i++)
                {
                    var text = columns[i].InnerText?.Trim();
                    if (!string.IsNullOrEmpty(text) && text.Length > 10 && text.Length < 200)
                    {
                        LogDebug($"Row {rowIndex}: Description: {text.Substring(0, Math.Min(50, text.Length))}...");
                        return text;
                    }
                }

                return "Quote item";
            }
            catch
            {
                return "Quote item";
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

                foreach (var column in columns)
                {
                    var html = column.InnerHtml;

                    var drawingMatch = Regex.Match(html, @"Drawing\s+(?:Number|No)[:\s]*([A-Z0-9\.\-]+)", RegexOptions.IgnoreCase);
                    if (drawingMatch.Success)
                    {
                        drawingNumber = drawingMatch.Groups[1].Value;
                    }

                    var revMatch = Regex.Match(html, @"rev[:\s\.]*([A-Z0-9]+)", RegexOptions.IgnoreCase);
                    if (revMatch.Success)
                    {
                        revision = revMatch.Groups[1].Value;
                    }
                }

                return (drawingNumber, revision);
            }
            catch
            {
                return ("", "00");
            }
        }

        /// <summary>
        /// 🔥 Extract lead time from columns
        /// </summary>
        private string ExtractLeadTime(HtmlNodeCollection columns, int rowIndex)
        {
            try
            {
                foreach (var column in columns)
                {
                    var text = column.InnerText;
                    var leadTimeMatch = Regex.Match(text, @"(\d+)\s*(?:days?|weeks?|weken?)", RegexOptions.IgnoreCase);
                    if (leadTimeMatch.Success)
                    {
                        var leadTime = leadTimeMatch.Groups[1].Value;
                        LogDebug($"Row {rowIndex}: Lead time: {leadTime} days");
                        return leadTime;
                    }
                }

                return "14"; // Default 14 days
            }
            catch
            {
                return "14";
            }
        }

        /// <summary>
        /// 🔥 ENHANCED: Extract quote prices using global data and table data
        /// </summary>
        private (string quotedPrice, string totalPrice) ExtractEnhancedQuotePrices(HtmlNodeCollection columns, int rowIndex,
            (string rfqNumber, string totalPrice, string quoteDate, string validUntil) globalQuoteInfo)
        {
            try
            {
                LogDebug($"Row {rowIndex}: Starting enhanced quote price extraction");

                // Method 1: Use global price info if available
                if (globalQuoteInfo.totalPrice != "0.00")
                {
                    LogDebug($"Row {rowIndex}: Using global quote price: {globalQuoteInfo.totalPrice}");
                    return (globalQuoteInfo.totalPrice, globalQuoteInfo.totalPrice);
                }

                // Method 2: Extract from table cells
                foreach (var column in columns)
                {
                    var text = column.InnerText;
                    var price = ExtractPriceFromText(text);
                    if (price != "0.00")
                    {
                        LogDebug($"Row {rowIndex}: Quote price found in table cell: {price}");
                        return (price, price);
                    }
                }

                LogDebug($"Row {rowIndex}: No quote prices found, using defaults");
                return ("0.00", "0.00");
            }
            catch (Exception ex)
            {
                LogDebug($"Row {rowIndex}: Error in quote price extraction: {ex.Message}");
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
                var pricePatterns = new[]
                {
                    @"€\s*(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})",
                    @"\$\s*(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})",
                    @"(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s*€",
                    @"(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s*\$",
                    @"(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})"
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
                            return cleanPrice;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogDebug($"Error extracting price from '{text}': {ex.Message}");
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

                // Handle European format
                if (price.Contains(",") && !price.Contains("."))
                {
                    price = price.Replace(",", ".");
                }
                else if (price.Contains(",") && price.Contains("."))
                {
                    var lastComma = price.LastIndexOf(',');
                    var lastDot = price.LastIndexOf('.');
                    if (lastComma > lastDot)
                    {
                        price = price.Replace(".", "").Replace(",", ".");
                    }
                }

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

        /// <summary>
        /// 🔥 Determine quote priority based on price
        /// </summary>
        private string DeterminePriority(string quotedPrice)
        {
            try
            {
                if (decimal.TryParse(quotedPrice, out decimal price))
                {
                    if (price > 10000) return "High";
                    if (price > 1000) return "Medium";
                }
                return "Normal";
            }
            catch
            {
                return "Normal";
            }
        }

        // ===============================================
        // BACKWARD COMPATIBILITY METHODS
        // ===============================================

        public List<QuoteLine> ExtractQuoteLinesOutlookForwardMail()
        {
            LogDebug("=== FALLBACK: Using enhanced extraction for Outlook forward mail ===");
            return ExtractQuoteLinesWeirDomain(); // Use enhanced method as fallback
        }
    }
}