// ===============================================
// COMPLETE FIXED OrderLogicService.cs
// Replace the entire file content with this
// ===============================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Mkg_Elcotec_Automation.Models;

namespace Mkg_Elcotec_Automation.Services
{
    public static class OrderLogicService
    {
        private static List<string> processingLog = new List<string>();

        // Get logs for debugging
        public static List<string> GetProcessingLog() => new List<string>(processingLog);
        public static void ClearProcessingLog() => processingLog.Clear();

        /// <summary>
        /// Main order extraction method with enhanced price extraction
        /// </summary>
        public static List<OrderLine> ExtractOrders(string subject, string emailBody, string emailDomain)
        {
            try
            {
                Console.WriteLine($"🎯 ExtractOrders called with domain: {emailDomain}");
                var orders = new List<OrderLine>();

                // Extract article codes from email content
                var articleCodes = ExtractArticleCodesInline(subject + " " + emailBody);
                Console.WriteLine($"📦 Found {articleCodes.Count} potential article codes: {string.Join(", ", articleCodes)}");

                if (!articleCodes.Any())
                {
                    // If no specific article codes found, create a generic order
                    articleCodes.Add("UNKNOWN-ARTICLE-" + DateTime.Now.ToString("MMddss"));
                    Console.WriteLine($"📦 No article codes found - created generic: {articleCodes[0]}");
                }

                foreach (var articleCode in articleCodes)
                {
                    // Extract order information
                    var orderInfo = ExtractOrderDetailsInline(subject, emailBody);

                    // ✅ ENHANCED: Extract price information using built-in price extraction
                    var (unitPrice, totalPrice) = ExtractPricesFromEmail(
                        subject,
                        emailBody,
                        emailDomain,
                        orderInfo.Quantity ?? "1"
                    );

                    Console.WriteLine($"💰 Extracted prices for {articleCode}: Unit={unitPrice}, Total={totalPrice}");

                    var order = new OrderLine
                    {
                        ArtiCode = articleCode,
                        Description = orderInfo.Description ?? $"Order for {articleCode}",
                        PoNumber = orderInfo.PoNumber ?? GeneratePoNumberInline(subject, articleCode),
                        Quantity = orderInfo.Quantity ?? "1",
                        Unit = orderInfo.Unit ?? "ST",
                        UnitPrice = unitPrice,  // ✅ Use extracted price instead of "0.00"
                        TotalPrice = totalPrice, // ✅ Use calculated total instead of "0.00"
                        OrderDate = DateTime.Now.ToString("yyyy-MM-dd"),
                        RequestedDeliveryDate = orderInfo.RequestedDelivery ?? "",
                        OrderStatus = "New",
                        Priority = DetermineOrderPriorityInline(subject, emailBody)
                    };

                    order.SetExtractionDetails("EMAIL_EXTRACTION", emailDomain);
                    orders.Add(order);

                    processingLog.Add($"✓ Created order: {articleCode} (PO: {order.PoNumber}, Price: €{unitPrice})");
                    Console.WriteLine($"   ✅ Created order: {articleCode} (PO: {order.PoNumber}, Price: €{unitPrice})");
                }

                processingLog.Add($"Extracted {orders.Count} orders from email");
                Console.WriteLine($"🎉 ExtractOrders completed: {orders.Count} orders created");
                return orders;
            }
            catch (Exception ex)
            {
                processingLog.Add($"Error in ExtractOrders: {ex.Message}");
                Console.WriteLine($"❌ ExtractOrders ERROR: {ex.Message}");
                return new List<OrderLine>();
            }
        }
        /// <summary>
        /// SIMPLE DEBUG: Just try to find any price in the email
        /// </summary>
        private static (string unitPrice, string totalPrice) ExtractPricesFromEmail(
            string subject,
            string body,
            string senderDomain,
            string quantity = "1")
        {
            Console.WriteLine($"🔍 === SIMPLE PRICE DEBUG ===");
            Console.WriteLine($"🔍 Domain: {senderDomain}");
            Console.WriteLine($"🔍 Subject: {subject?.Substring(0, Math.Min(subject?.Length ?? 0, 100))}...");

            try
            {
                // Combine subject and body for searching
                var fullText = (subject ?? "") + " " + (body ?? "");
                Console.WriteLine($"🔍 Full text length: {fullText.Length}");

                // Look for any Euro amounts
                var euroPattern = @"€\s*(\d+[.,]\d{2})";
                var euroMatches = Regex.Matches(fullText, euroPattern, RegexOptions.IgnoreCase);
                Console.WriteLine($"🔍 Found {euroMatches.Count} Euro patterns");

                foreach (Match match in euroMatches)
                {
                    var priceStr = match.Groups[1].Value;
                    Console.WriteLine($"🔍 Euro candidate: {priceStr}");

                    var cleanPrice = CleanPriceSimple(priceStr);
                    if (cleanPrice != "0.00")
                    {
                        Console.WriteLine($"✅ Using Euro price: {cleanPrice}");
                        return (cleanPrice, cleanPrice);
                    }
                }

                // Look for any decimal numbers that could be prices
                var decimalPattern = @"\b(\d+[.,]\d{2})\b";
                var decimalMatches = Regex.Matches(fullText, decimalPattern);
                Console.WriteLine($"🔍 Found {decimalMatches.Count} decimal patterns");

                foreach (Match match in decimalMatches.Cast<Match>()) // Only check first 5
                {
                    var numberStr = match.Groups[1].Value;
                    Console.WriteLine($"🔍 Decimal candidate: {numberStr}");

                    var cleanPrice = CleanPriceSimple(numberStr);
                    if (cleanPrice != "0.00")
                    {
                        if (decimal.TryParse(cleanPrice, out decimal price) && price >= 1.00m && price <= 10000.00m)
                        {
                            Console.WriteLine($"✅ Using decimal price: {cleanPrice}");
                            return (cleanPrice, cleanPrice);
                        }
                    }
                }

                Console.WriteLine($"❌ No prices found anywhere");
                return ("0.00", "0.00");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in price extraction: {ex.Message}");
                return ("0.00", "0.00");
            }
        }

        /// <summary>
        /// Simple price cleaning
        /// </summary>
        private static string CleanPriceSimple(string price)
        {
            if (string.IsNullOrEmpty(price))
                return "0.00";

            try
            {
                // Remove common non-numeric characters
                price = price.Replace("€", "").Replace("$", "").Replace("£", "").Trim();

                // Handle European format (comma as decimal)
                if (price.Contains(","))
                {
                    price = price.Replace(",", ".");
                }

                // Try to parse
                if (decimal.TryParse(price, out decimal result) && result > 0)
                {
                    return result.ToString("F2");
                }

                return "0.00";
            }
            catch
            {
                return "0.00";
            }
        }
         
        private static (string unitPrice, string totalPrice) ExtractPricesAggressively(string text, string quantity)
        {
            if (string.IsNullOrEmpty(text))
                return ("0.00", "0.00");

            try
            {
                Console.WriteLine($"🔍 Starting aggressive price detection...");

                // Find all decimal numbers in the text
                var allDecimals = new List<decimal>();
                var decimalPattern = @"\b(\d{1,6}[.,]\d{2})\b";
                var matches = Regex.Matches(text, decimalPattern);

                Console.WriteLine($"🔍 Found {matches.Count} potential decimal numbers");

                foreach (Match match in matches)
                {
                    var numberStr = match.Groups[1].Value;
                    var cleanNumber = CleanAndValidatePrice(numberStr);

                    if (cleanNumber != "0.00" && decimal.TryParse(cleanNumber, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal number))
                    {
                        // Filter out values that are unlikely to be prices
                        if (number >= 0.01m && number <= 999999.99m)
                        {
                            allDecimals.Add(number);
                            Console.WriteLine($"🔍 Potential price candidate: €{number:F2}");
                        }
                    }
                }

                if (allDecimals.Any())
                {
                    // Use heuristics to pick the most likely price
                    var mostLikelyPrice = allDecimals
                        .Where(d => d >= 1.00m && d <= 50000.00m)  // Reasonable price range
                        .OrderBy(d => Math.Abs(d - 100m))          // Prefer prices around €100 (common range)
                        .FirstOrDefault();

                    if (mostLikelyPrice > 0)
                    {
                        var priceStr = mostLikelyPrice.ToString("F2", CultureInfo.InvariantCulture);
                        Console.WriteLine($"✅ Selected most likely price: €{priceStr}");

                        // Calculate total
                        var totalPrice = priceStr;
                        if (decimal.TryParse(quantity, out decimal qty) && qty > 1)
                        {
                            totalPrice = (mostLikelyPrice * qty).ToString("F2", CultureInfo.InvariantCulture);
                        }

                        return (priceStr, totalPrice);
                    }
                }

                Console.WriteLine($"❌ No reasonable price candidates found");
                return ("0.00", "0.00");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in aggressive price extraction: {ex.Message}");
                return ("0.00", "0.00");
            }
        }
        /// <summary>
        /// Extract prices from structured email content (tables, lists)
        /// </summary>
        private static (string unitPrice, string totalPrice) ExtractPricesFromStructuredContent(string content, string quantity)
        {
            if (string.IsNullOrEmpty(content))
                return ("0.00", "0.00");

            try
            {
                Console.WriteLine($"🔍 Analyzing structured content for prices...");

                // Look for common price table structures in Weir emails
                var structuredPatterns = new[]
                {
            @"Unit\s*Price[:\s]*([€$£]?\s*\d+[.,]\d{2})",                    // Unit Price: €123.45
            @"Price[:\s]*([€$£]?\s*\d+[.,]\d{2})",                          // Price: €123.45
            @"Total[:\s]*([€$£]?\s*\d+[.,]\d{2})",                          // Total: €123.45
            @"Amount[:\s]*([€$£]?\s*\d+[.,]\d{2})",                         // Amount: €123.45
            @"Cost[:\s]*([€$£]?\s*\d+[.,]\d{2})",                           // Cost: €123.45
            @"\|\s*([€$£]?\s*\d+[.,]\d{2})\s*\|",                          // | €123.45 |
            @"([€$£]\s*\d+[.,]\d{2})\s*\|\s*Qty",                          // €123.45 | Qty
            @"Qty:\s*\d+\s*\|\s*([€$£]?\s*\d+[.,]\d{2})",                  // Qty: 1 | €123.45
        };

                foreach (var pattern in structuredPatterns)
                {
                    var matches = Regex.Matches(content, pattern, RegexOptions.IgnoreCase);
                    foreach (Match match in matches)
                    {
                        var priceStr = match.Groups[1].Value;
                        var cleanPrice = CleanAndValidatePrice(priceStr);

                        if (cleanPrice != "0.00")
                        {
                            Console.WriteLine($"✅ Found structured price with pattern '{pattern}': {cleanPrice}");

                            // Calculate total if we have quantity
                            var totalPrice = cleanPrice;
                            if (decimal.TryParse(quantity, out decimal qty) && qty > 1)
                            {
                                if (decimal.TryParse(cleanPrice, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal unitPriceDecimal))
                                {
                                    totalPrice = (unitPriceDecimal * qty).ToString("F2", CultureInfo.InvariantCulture);
                                }
                            }

                            return (cleanPrice, totalPrice);
                        }
                    }
                }

                Console.WriteLine($"❌ No structured prices found");
                return ("0.00", "0.00");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in structured price extraction: {ex.Message}");
                return ("0.00", "0.00");
            }
        }
        /// <summary>
        /// Extract prices from Weir emails
        /// </summary>
        private static (string unitPrice, string totalPrice) ExtractWeirPrices(
            string subject,
            string body,
            string quantity)
        {
            try
            {
                // Method 1: Extract from plain text patterns
                var textPrices = ExtractPricesFromText(body + " " + subject, quantity);
                if (textPrices.unitPrice != "0.00")
                {
                    Console.WriteLine($"✅ Found Weir prices in text: Unit={textPrices.unitPrice}, Total={textPrices.totalPrice}");
                    return textPrices;
                }

                Console.WriteLine("❌ No prices found in Weir email");
                return ("0.00", "0.00");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in Weir price extraction: {ex.Message}");
                return ("0.00", "0.00");
            }
        }

        /// <summary>
        /// Extract prices from plain text using regex patterns
        /// </summary>
        private static (string unitPrice, string totalPrice) ExtractPricesFromText(string text, string quantity)
        {
            if (string.IsNullOrEmpty(text))
                return ("0.00", "0.00");

            try
            {
                // Multiple price patterns to try
                var pricePatterns = new[]
                {
                    @"(?:price|cost|amount|total)[:\s]*€\s*(\d+[.,]\d{2})",           // Price: €123.45
                    @"(?:price|cost|amount|total)[:\s]*\$\s*(\d+[.,]\d{2})",          // Price: $123.45
                    @"€\s*(\d+[.,]\d{2})",                                            // €123.45
                    @"\$\s*(\d+[.,]\d{2})",                                           // $123.45
                    @"(\d+[.,]\d{2})\s*€",                                            // 123.45 €
                    @"(\d+[.,]\d{2})\s*\$",                                           // 123.45 $
                    @"(\d+[.,]\d{2})\s*EUR",                                          // 123.45 EUR
                    @"(\d+[.,]\d{2})\s*USD",                                          // 123.45 USD
                    @"(?:unit price|each)[:\s]*(\d+[.,]\d{2})",                       // Unit price: 123.45
                    @"(?:total)[:\s]*(\d+[.,]\d{2})"                                  // Total: 123.45
                };

                foreach (var pattern in pricePatterns)
                {
                    var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
                    foreach (Match match in matches)
                    {
                        var priceStr = match.Groups[1].Value;
                        var cleanPrice = CleanAndValidatePrice(priceStr);

                        if (cleanPrice != "0.00")
                        {
                            Console.WriteLine($"✅ Found price with pattern '{pattern}': {cleanPrice}");

                            // Calculate total if we have quantity
                            var totalPrice = cleanPrice;
                            if (decimal.TryParse(quantity, out decimal qty) && qty > 1)
                            {
                                if (decimal.TryParse(cleanPrice, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal unitPriceDecimal))
                                {
                                    totalPrice = (unitPriceDecimal * qty).ToString("F2", CultureInfo.InvariantCulture);
                                }
                            }

                            return (cleanPrice, totalPrice);
                        }
                    }
                }

                return ("0.00", "0.00");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error extracting from text: {ex.Message}");
                return ("0.00", "0.00");
            }
        }

        /// <summary>
        /// Generic price extraction for other email types
        /// </summary>
        private static (string unitPrice, string totalPrice) ExtractGenericPrices(
            string subject,
            string body,
            string quantity)
        {
            Console.WriteLine("🔧 Using generic price extraction");
            return ExtractPricesFromText(body + " " + subject, quantity);
        }

        /// <summary>
        /// Clean and validate price string
        /// </summary>
        private static string CleanAndValidatePrice(string price)
        {
            if (string.IsNullOrWhiteSpace(price))
                return "0.00";

            try
            {
                // Remove currency symbols and extra whitespace
                price = price.Replace("€", "").Replace("$", "").Replace("£", "")
                            .Replace("USD", "").Replace("EUR", "").Replace("GBP", "")
                            .Trim();

                // Handle European decimal format (comma as decimal separator)
                if (price.Contains(",") && !price.Contains("."))
                {
                    price = price.Replace(",", ".");
                }
                else if (price.Contains(",") && price.Contains("."))
                {
                    // European format: 1.234,56 -> 1234.56
                    var lastComma = price.LastIndexOf(',');
                    var lastDot = price.LastIndexOf('.');
                    if (lastComma > lastDot)
                    {
                        price = price.Replace(".", "").Replace(",", ".");
                    }
                }

                // Remove any remaining non-numeric characters except decimal point
                price = Regex.Replace(price, @"[^\d\.]", "");

                // Validate and format
                if (decimal.TryParse(price, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result) && result > 0)
                {
                    return result.ToString("F2", CultureInfo.InvariantCulture);
                }

                return "0.00";
            }
            catch
            {
                return "0.00";
            }
        }

        // ===============================================
        // INLINE HELPER METHODS
        // ===============================================

        /// <summary>
        /// Extract article codes from text
        /// </summary>
        private static List<string> ExtractArticleCodesInline(string text)
        {
            var articleCodes = new HashSet<string>();

            if (string.IsNullOrEmpty(text))
                return articleCodes.ToList();

            try
            {
                // Enhanced patterns for article codes
                var patterns = new[]
                {
            @"\b(\d{3}\.\d{3}\.\d{3})\b",                    // 815.920.098
            @"\b(\d{3}\.\d{3}\.\d{2,3})\b",                  // 816.940.393
            @"\b(\d{3}\.\d{3}\.\d{1,3}[A-Z]*)\b",           // 870.001.272A
            @"\b([A-Z]{2,4}\d{3,4})\b",                      // RAL3020, NL002
            @"\b(\d{3}\.\d{2,3}\.\d{2,3})\b",               // 500.143.527
            @"\b([A-Z]{3}\d{3})\b",                          // FFC000
            @"\b(\d{3}-\d{2}-\d{3})\b",                      // 475-25-041
            @"\b(\d{3}\.\d{3}\.\d{4}[A-Z]*)\b",             // 897.010.1478
            @"(?:article|part|item)[:\s#]*([A-Z0-9\.\-]{6,})", // Article: ABC123
            @"(?:sku|pn|part\s*number)[:\s#]*([A-Z0-9\.\-]{6,})", // SKU: ABC123
        };

                foreach (var pattern in patterns)
                {
                    var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
                    foreach (Match match in matches)
                    {
                        var code = match.Groups[1].Value.ToUpper().Trim();
                        if (code.Length >= 3 && code.Length <= 20)
                        {
                            articleCodes.Add(code);
                            Console.WriteLine($"📦 Found article code: {code}");
                        }
                    }
                }

                // If no article codes found with patterns, try to extract from email structure
                if (!articleCodes.Any())
                {
                    var structuredCodes = ExtractCodesFromStructuredContent(text);
                    foreach (var code in structuredCodes)
                    {
                        articleCodes.Add(code);
                    }
                }

                return articleCodes.ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error extracting article codes: {ex.Message}");
                return new List<string>();
            }
        }
        private static List<string> ExtractCodesFromStructuredContent(string text)
        {
            var codes = new List<string>();

            try
            {
                // Look for structured patterns like:
                // - ARTICLE-123 | Qty: 1 | PO: PO-456
                // - Item: ABC.DEF.GHI | Quantity: 2

                var structuredPatterns = new[]
                {
            @"-\s*([A-Z0-9\.\-]{3,15})\s*\|",               // - CODE |
            @"item[:\s]*([A-Z0-9\.\-]{3,15})",              // Item: CODE
            @"article[:\s]*([A-Z0-9\.\-]{3,15})",           // Article: CODE
            @"code[:\s]*([A-Z0-9\.\-]{3,15})",              // Code: CODE
            @"([A-Z0-9\.\-]{3,15})\s*\|\s*qty",            // CODE | Qty:
        };

                foreach (var pattern in structuredPatterns)
                {
                    var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
                    foreach (Match match in matches)
                    {
                        var code = match.Groups[1].Value.ToUpper().Trim();
                        if (IsValidArticleCode(code))
                        {
                            codes.Add(code);
                            Console.WriteLine($"📦 Found structured article code: {code}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error extracting structured codes: {ex.Message}");
            }

            return codes;
        }
        private static bool IsValidArticleCode(string code)
        {
            if (string.IsNullOrEmpty(code) || code.Length < 3 || code.Length > 20)
                return false;

            // Must contain at least one digit or one letter
            if (!Regex.IsMatch(code, @"[A-Z0-9]"))
                return false;

            // Exclude common false positives
            var excludePatterns = new[]{
                                        @"^\d{1,2}$",           // Single/double digits
                                        @"^[A-Z]{1,2}$",        // Single/double letters
                                        @"^(THE|AND|FOR|WITH)$", // Common words
                                    };

            foreach (var exclude in excludePatterns)
            {
                if (Regex.IsMatch(code, exclude, RegexOptions.IgnoreCase))
                    return false;
            }

            return true;
        }

        /// <summary>
        /// Extract detailed order information from email content
        /// </summary>
        private static (string PoNumber, string Description, string Quantity, string Unit, string RequestedDelivery) ExtractOrderDetailsInline(string subject, string body)
        {
            // Extract PO number
            string poNumber = null;
            var poPattern = @"(po|order)[:\-\s#]*(\d+)";
            var poMatch = Regex.Match(subject + " " + body, poPattern, RegexOptions.IgnoreCase);
            if (poMatch.Success)
            {
                poNumber = "PO-" + poMatch.Groups[2].Value;
            }

            // Extract quantity
            var qtyPattern = @"(qty|quantity|aantal)[:\-\s]*(\d+)";
            var qtyMatch = Regex.Match(subject + " " + body, qtyPattern, RegexOptions.IgnoreCase);
            string quantity = qtyMatch.Success ? qtyMatch.Groups[2].Value : "1";

            // Extract unit
            var unitPattern = @"(\d+)\s*(st|pcs|pieces|stuks|each)";
            var unitMatch = Regex.Match(body, unitPattern, RegexOptions.IgnoreCase);
            string unit = unitMatch.Success ? unitMatch.Groups[2].Value.ToUpper() : "ST";

            // Extract description (use subject if no specific description found)
            string description = subject?.Length > 10 ? subject.Substring(0, Math.Min(subject.Length, 100)) : "Order item";

            // Extract delivery date
            var deliveryPattern = @"deliver[y]?[:\-\s]*([\d\-\/\.]+)";
            var deliveryMatch = Regex.Match(body, deliveryPattern, RegexOptions.IgnoreCase);
            string requestedDelivery = "";
            if (deliveryMatch.Success)
            {
                if (DateTime.TryParse(deliveryMatch.Groups[1].Value, out DateTime deliveryDate))
                {
                    requestedDelivery = deliveryDate.ToString("yyyy-MM-dd");
                }
            }

            return (poNumber, description, quantity, unit, requestedDelivery);
        }

        /// <summary>
        /// Generate PO number if not found
        /// </summary>
        private static string GeneratePoNumberInline(string subject, string articleCode)
        {
            var timestamp = DateTime.Now.ToString("yyyyMMddHHmm");
            var shortArticle = articleCode.Length > 8 ? articleCode.Substring(0, 8) : articleCode;
            return $"PO-{shortArticle}-{timestamp}";
        }

        /// <summary>
        /// Determine order priority based on content
        /// </summary>
        private static string DetermineOrderPriorityInline(string subject, string body)
        {
            var urgentKeywords = new[] { "urgent", "asap", "rush", "emergency", "critical", "urgent", "spoed" };
            var text = (subject + " " + body).ToLower();

            foreach (var keyword in urgentKeywords)
            {
                if (text.Contains(keyword))
                {
                    return "High";
                }
            }

            return "Normal";
        }

        /// <summary>
        /// Enhanced order extraction that tries to find multiple orders in one email
        /// </summary>
        public static List<OrderLine> ExtractOrdersSafe(string subject, string emailBody, string emailDomain)
        {
            try
            {
                return ExtractOrders(subject, emailBody, emailDomain);
            }
            catch (Exception ex)
            {
                processingLog.Add($"Safe extraction failed: {ex.Message}");
                Console.WriteLine($"❌ ExtractOrdersSafe ERROR: {ex.Message}");
                return new List<OrderLine>();
            }
        }

        /// <summary>
        /// Validate if text contains order-like content
        /// </summary>
        public static bool ContainsOrderContent(string subject, string body)
        {
            if (string.IsNullOrEmpty(subject) && string.IsNullOrEmpty(body))
                return false;

            var orderKeywords = new[]
            {
                "purchase order", "po #", "order", "bestelling", "inkooporder",
                "aanvraag", "bestellung", "commande"
            };

            var text = (subject + " " + body).ToLower();
            return orderKeywords.Any(keyword => text.Contains(keyword));
        }

        /// <summary>
        /// Enhanced processing for complex orders
        /// </summary>
        public static List<OrderLine> ProcessRawOrders(List<OrderLine> rawOrders, string customerDomain)
        {
            var processedOrders = new List<OrderLine>();

            try
            {
                foreach (var order in rawOrders)
                {
                    try
                    {
                        // Apply basic validation and enhancement
                        if (!string.IsNullOrEmpty(order.ArtiCode))
                        {
                            // Enhance with customer domain info
                            order.ExtractionDomain = customerDomain;

                            // Add to processed list
                            processedOrders.Add(order);
                        }
                    }
                    catch (Exception ex)
                    {
                        processingLog.Add($"Error processing order {order.ArtiCode}: {ex.Message}");
                    }
                }

                processingLog.Add($"Processed {processedOrders.Count}/{rawOrders.Count} orders successfully");
                return processedOrders;
            }
            catch (Exception ex)
            {
                processingLog.Add($"Critical error in order processing: {ex.Message}");
                return new List<OrderLine>();
            }
        }

        /// <summary>
        /// Extract order details (supports both 3 and 4 parameter calls)
        /// </summary>
        public static (string PoNumber, string Description, string Quantity, string Unit, string RequestedDelivery) ExtractOrderDetails(string subject, string body)
        {
            return ExtractOrderDetailsInline(subject, body);
        }

        /// <summary>
        /// Overload for backward compatibility
        /// </summary>
        public static (string PoNumber, string Description, string Quantity, string Unit, string RequestedDelivery) ExtractOrderDetails(string subject, string body, string domain)
        {
            return ExtractOrderDetailsInline(subject, body);
        }

        /// <summary>
        /// Overload for backward compatibility with 4 parameters
        /// </summary>
        public static (string PoNumber, string Description, string Quantity, string Unit, string RequestedDelivery) ExtractOrderDetails(string subject, string body, string domain, string additionalInfo)
        {
            return ExtractOrderDetailsInline(subject, body);
        }

        /// <summary>
        /// Async wrapper for backward compatibility with EmailWorkflowService
        /// </summary>
        public static async Task<List<OrderLine>> ExtractOrders(string emailBody, string emailDomain, string subject, object attachments)
        {
            // Convert to synchronous call with correct parameter order
            await Task.Delay(1); // Minimal async operation
            return ExtractOrders(subject, emailBody, emailDomain);
        }
    }
}