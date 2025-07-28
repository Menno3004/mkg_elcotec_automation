// ===============================================
// FIXED OrderLogicService.cs - Enhanced Price Extraction
// This integrates with the existing system while adding price extraction
// ===============================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Mkg_Elcotec_Automation.Models;
using Mkg_Elcotec_Automation.Utilities; // For ArticleCodeExtractor

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

                // Use ArticleCodeExtractor from Utilities
                var articleCodes = ArticleCodeExtractor.ExtractArticleCodes(subject + " " + emailBody);
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

                    // 🔥 ENHANCED: Extract price information with debugging
                    var (unitPrice, totalPrice) = ExtractPricesFromEmailEnhanced(
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
                        Unit = orderInfo.Unit ?? "PCS",
                        UnitPrice = unitPrice,   // 🔥 Use extracted price
                        TotalPrice = totalPrice, // 🔥 Use extracted total
                        RequestedDeliveryDate = orderInfo.RequestedDelivery ?? DateTime.Now.AddDays(14).ToString("yyyy-MM-dd"),
                        ExtractionMethod = "EMAIL_EXTRACTION",
                        EmailDomain = emailDomain
                    };

                    orders.Add(order);
                    Console.WriteLine($"✅ Created order for {articleCode} with price {unitPrice}");
                }

                Console.WriteLine($"📦 Total orders created: {orders.Count}");
                return orders;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in ExtractOrders: {ex.Message}");
                return new List<OrderLine>();
            }
        }

        /// <summary>
        /// 🔥 ENHANCED price extraction with detailed debugging and Weir-specific patterns
        /// </summary>
        private static (string unitPrice, string totalPrice) ExtractPricesFromEmailEnhanced(string subject, string body, string domain, string quantity)
        {
            Console.WriteLine($"🔍 === ENHANCED PRICE EXTRACTION DEBUG ===");
            Console.WriteLine($"🔍 Domain: {domain}");
            Console.WriteLine($"🔍 Quantity: {quantity}");

            try
            {
                string fullText = (subject + " " + body).Replace("\r\n", " ").Replace("\n", " ");
                Console.WriteLine($"🔍 Text preview: {fullText.Substring(0, Math.Min(200, fullText.Length))}...");

                // 🎯 WEIR-SPECIFIC PRICE PATTERNS (Enhanced)
                if (domain.Contains("weir") || domain.Contains("coupahost"))
                {
                    var weirPatterns = new[]
                    {
                        @"PO\s+Total[:\s]*(\d+[.,]\d{3})[.,](\d{2})\s*EUR",   // "PO Total: 35.520,00 EUR"
                        @"Total[:\s]*(\d+[.,]\d{3})[.,](\d{2})\s*EUR",        // "Total: 35.520,00 EUR"
                        @"(\d+[.,]\d{3})[.,](\d{2})\s*EUR",                   // "35.520,00 EUR"
                        @"PO\s+Total[:\s]*(\d+)[.,](\d{2})\s*EUR",            // "PO Total: 980,00 EUR"
                        @"Total[:\s]*(\d+)[.,](\d{2})\s*EUR",                 // "Total: 980,00 EUR"
                        @"(\d+)[.,](\d{2})\s*EUR",                            // "980,00 EUR"
                        @"€\s*(\d+[.,]\d{3})[.,](\d{2})",                     // "€ 35.520,00"
                        @"€\s*(\d+)[.,](\d{2})",                              // "€ 980,00"
                    };

                    foreach (var pattern in weirPatterns)
                    {
                        var match = Regex.Match(fullText, pattern, RegexOptions.IgnoreCase);
                        if (match.Success)
                        {
                            Console.WriteLine($"🎯 WEIR PATTERN MATCHED: {pattern}");

                            string priceStr;
                            if (match.Groups.Count >= 3 && !string.IsNullOrEmpty(match.Groups[2].Value))
                            {
                                // Format with thousands separator: 35.520,00 → 35520.00
                                priceStr = match.Groups[1].Value.Replace(",", "").Replace(".", "") + "." + match.Groups[2].Value;
                            }
                            else
                            {
                                // Simple format: 980,00 → 980.00
                                priceStr = match.Groups[1].Value.Replace(",", ".");
                            }

                            Console.WriteLine($"🎯 Reconstructed price: {priceStr}");

                            if (decimal.TryParse(priceStr, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal totalPrice))
                            {
                                var totalPriceFormatted = totalPrice.ToString("F2", CultureInfo.InvariantCulture);

                                // Calculate unit price
                                decimal unitPriceDecimal = totalPrice;
                                if (decimal.TryParse(quantity, out decimal qty) && qty > 0)
                                {
                                    unitPriceDecimal = totalPrice / qty;
                                }
                                var unitPriceFormatted = unitPriceDecimal.ToString("F2", CultureInfo.InvariantCulture);

                                Console.WriteLine($"✅ WEIR PRICE SUCCESS: Total={totalPriceFormatted}, Unit={unitPriceFormatted}");
                                return (unitPriceFormatted, totalPriceFormatted);
                            }
                        }
                    }
                    Console.WriteLine($"❌ No Weir-specific patterns matched");
                }

                // 🎯 GENERIC PRICE PATTERNS
                var genericPatterns = new[]
                {
                    @"(?:price|cost|amount|total)[:\s]*€\s*(\d+[.,]\d{2})",
                    @"(?:price|cost|amount|total)[:\s]*\$\s*(\d+[.,]\d{2})",
                    @"€\s*(\d+[.,]\d{2})",
                    @"\$\s*(\d+[.,]\d{2})",
                    @"(\d+[.,]\d{2})\s*€",
                    @"(\d+[.,]\d{2})\s*\$"
                };

                foreach (var pattern in genericPatterns)
                {
                    var match = Regex.Match(fullText, pattern, RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        var priceStr = match.Groups[1].Value.Replace(",", ".");
                        if (decimal.TryParse(priceStr, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal price))
                        {
                            var unitPrice = price.ToString("F2", CultureInfo.InvariantCulture);
                            var totalPrice = CalculateTotalPrice(unitPrice, quantity);
                            Console.WriteLine($"✅ GENERIC PRICE SUCCESS: Unit={unitPrice}, Total={totalPrice}");
                            return (unitPrice, totalPrice);
                        }
                    }
                }

                Console.WriteLine($"❌ NO PRICES FOUND ANYWHERE");
                return ("0.00", "0.00");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in enhanced price extraction: {ex.Message}");
                return ("0.00", "0.00");
            }
        }

        // ===============================================
        // EXISTING HELPER METHODS - ENHANCED
        // ===============================================

        private static (string PoNumber, string Description, string Quantity, string Unit, string RequestedDelivery) ExtractOrderDetailsInline(string subject, string emailBody)
        {
            Console.WriteLine($"🔍 EXTRACTING ORDER DETAILS:");

            // IMPROVED PO NUMBER EXTRACTION
            string poNumber = ExtractPoNumberImproved(subject, emailBody);
            Console.WriteLine($"   PO Number: {poNumber}");

            // Enhanced quantity extraction - looks for patterns like "3 x 891.029.1541"
            var qtyPattern = @"(?:qty|quantity|aantal|line\s*\d+)[:\-\s]*(\d+)\s*x|(\d+)\s*(?:st|pcs|pieces|stuks|each)|(\d+)\s*x\s*\d+";
            var qtyMatch = Regex.Match(subject + " " + emailBody, qtyPattern, RegexOptions.IgnoreCase);
            string quantity = "1"; // default
            if (qtyMatch.Success)
            {
                quantity = qtyMatch.Groups[1].Value;
                if (string.IsNullOrEmpty(quantity)) quantity = qtyMatch.Groups[2].Value;
                if (string.IsNullOrEmpty(quantity)) quantity = qtyMatch.Groups[3].Value;
                if (string.IsNullOrEmpty(quantity)) quantity = "1";
            }
            Console.WriteLine($"   Quantity: {quantity}");

            // Enhanced unit extraction
            var unitPattern = @"(\d+)\s*(st|pcs|pieces|stuks|each)";
            var unitMatch = Regex.Match(emailBody, unitPattern, RegexOptions.IgnoreCase);
            string unit = unitMatch.Success ? unitMatch.Groups[2].Value.ToUpper() : "PCS";
            Console.WriteLine($"   Unit: {unit}");

            // Enhanced description extraction
            string description = "Order item";
            if (!string.IsNullOrEmpty(subject) && subject.Length > 10)
            {
                // Clean up subject for description
                var cleanSubject = subject.Replace("RE:", "").Replace("FW:", "").Replace("[External]", "").Trim();
                description = cleanSubject.Length > 100 ? cleanSubject.Substring(0, 100) : cleanSubject;
            }
            Console.WriteLine($"   Description: {description}");

            // Enhanced delivery date extraction
            var deliveryPattern = @"deliver[y]?[:\-\s]*([\d\-\/\.]+)|leverdatum[:\-\s]*([\d\-\/\.]+)|delivery\s+date[:\-\s]*([\d\-\/\.]+)";
            var deliveryMatch = Regex.Match(emailBody, deliveryPattern, RegexOptions.IgnoreCase);
            string requestedDelivery = "";
            if (deliveryMatch.Success)
            {
                var dateStr = deliveryMatch.Groups[1].Value;
                if (string.IsNullOrEmpty(dateStr)) dateStr = deliveryMatch.Groups[2].Value;
                if (string.IsNullOrEmpty(dateStr)) dateStr = deliveryMatch.Groups[3].Value;

                if (DateTime.TryParse(dateStr, out DateTime deliveryDate))
                {
                    requestedDelivery = deliveryDate.ToString("yyyy-MM-dd");
                }
            }
            Console.WriteLine($"   Delivery Date: {requestedDelivery}");

            return (poNumber, description, quantity, unit, requestedDelivery);
        }

        private static string ExtractPoNumberImproved(string subject, string body)
        {
            var text = subject + " " + body;

            // Priority 1: Weir Purchase Order patterns
            var weirPatterns = new[]
            {
                @"Purchase\s+Order\s+#(\d{10})",                    // "Purchase Order #4501533672"
                @"Purchase\s+Order\s+(\d{10})",
                @"inkooporder\s+(\d{10})",
                @"PO[\s#:]*(\d{10})",
                @"PO[\s#:]*(\d{8,12})",
            };

            // Try Weir-specific patterns first
            foreach (var pattern in weirPatterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    return $"PO-{match.Groups[1].Value}";
                }
            }

            // Fallback: Generate from timestamp
            return $"PO-AUTO-{DateTime.Now:yyyyMMddHHmm}";
        }

        private static string GeneratePoNumberInline(string subject, string articleCode)
        {
            // First try to extract from subject
            var extractedPo = ExtractPoNumberImproved(subject, "");
            if (!extractedPo.Contains("AUTO"))
            {
                return extractedPo;
            }

            // Fallback generation
            var timestamp = DateTime.Now.ToString("yyyyMMddHHmm");
            var shortArticle = articleCode.Length > 8 ? articleCode.Substring(0, 8).Replace(".", "") : articleCode.Replace(".", "");
            return $"PO-{shortArticle}-{timestamp}";
        }

        private static string CalculateTotalPrice(string unitPrice, string quantity)
        {
            try
            {
                if (decimal.TryParse(unitPrice, out decimal unit) &&
                    decimal.TryParse(quantity, out decimal qty))
                {
                    var total = unit * qty;
                    return total.ToString("F2", CultureInfo.InvariantCulture);
                }
                return unitPrice; // If calculation fails, return unit price
            }
            catch
            {
                return unitPrice;
            }
        }

        // ===============================================
        // ASYNC COMPATIBILITY METHODS
        // ===============================================

        public static async Task<List<OrderLine>> ExtractOrders(string emailBody, string emailDomain, string subject, object attachments)
        {
            try
            {
                await Task.Delay(1);
                return ExtractOrders(subject, emailBody, emailDomain);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in ExtractOrders (4-param): {ex.Message}");
                return new List<OrderLine>();
            }
        }
    }
}