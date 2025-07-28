// ===============================================
// UPDATED QuoteLogicService.cs - Using ArticleCodeExtractor
// MINIMAL CHANGES: Only replaced article code extraction with ArticleCodeExtractor
// ===============================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Mkg_Elcotec_Automation.Models;
using Mkg_Elcotec_Automation.Utilities; // 🆕 ADD THIS USING

namespace Mkg_Elcotec_Automation.Services
{
    public static class QuoteLogicService
    {
        private static List<string> processingLog = new List<string>();

        #region Public Properties
        public static List<string> GetProcessingLog() => new List<string>(processingLog);
        public static void ClearProcessingLog() => processingLog.Clear();
        #endregion

        #region Core Quote Business Logic

        /// <summary>
        /// Process and validate quotes with business rules applied
        /// Input: Raw quotes from EmailContentAnalyzer
        /// Output: Business-validated and processed quotes ready for MKG injection
        /// </summary>
        public static List<QuoteLine> ProcessQuotesForMkgInjection(List<QuoteLine> rawQuotes, string customerDomain = null)
        {
            try
            {
                processingLog.Add($"=== PROCESSING {rawQuotes.Count} QUOTES FOR MKG INJECTION ===");

                var processedQuotes = new List<QuoteLine>();

                foreach (var quote in rawQuotes)
                {
                    try
                    {
                        // Apply business validation
                        if (!ValidateQuoteBusinessRules(quote))
                        {
                            processingLog.Add($"Quote validation failed for ArtiCode: {quote.ArtiCode}");
                            continue;
                        }

                        // Apply business processing rules
                        ApplyQuoteBusinessRules(quote, customerDomain);

                        // Calculate pricing
                        CalculateQuotePricing(quote);

                        // Assign priority based on business rules
                        AssignPriorityBasedOnBusinessRules(quote, customerDomain);

                        // Format for MKG standards
                        FormatQuoteForMkgStandards(quote);

                        processedQuotes.Add(quote);
                        processingLog.Add($"✓ Successfully processed quote: {quote.ArtiCode} (RFQ: {quote.RfqNumber})");
                    }
                    catch (Exception ex)
                    {
                        processingLog.Add($"Error processing quote {quote.ArtiCode}: {ex.Message}");
                    }
                }

                processingLog.Add($"Processed {processedQuotes.Count}/{rawQuotes.Count} quotes successfully");
                return processedQuotes;
            }
            catch (Exception ex)
            {
                processingLog.Add($"Critical error in quote processing: {ex.Message}");
                return new List<QuoteLine>();
            }
        }

        /// <summary>
        /// 🆕 ADDED: Extract quotes from email content using centralized article code extraction
        /// </summary>
        public static List<QuoteLine> ExtractQuotes(string subject, string emailBody, string emailDomain)
        {
            try
            {
                processingLog.Add($"🎯 ExtractQuotes called with domain: {emailDomain}");
                var quotes = new List<QuoteLine>();

                // 🆕 USE CENTRALIZED EXTRACTOR
                var articleCodes = ArticleCodeExtractor.ExtractArticleCodes(subject + " " + emailBody);
                processingLog.Add($"📦 Found {articleCodes.Count} valid article codes: {string.Join(", ", articleCodes)}");

                if (!articleCodes.Any())
                {
                    processingLog.Add("📦 No article codes found for quote extraction");
                    return quotes;
                }

                foreach (var articleCode in articleCodes)
                {
                    // Extract quote information
                    var quoteInfo = ExtractQuoteDetailsInline(subject, emailBody);

                    var quote = new QuoteLine
                    {
                        ArtiCode = articleCode,
                        Description = quoteInfo.Description ?? $"Quote for {articleCode}",
                        RfqNumber = quoteInfo.RfqNumber ?? GenerateRfqNumberInline(subject, articleCode),
                        Quantity = quoteInfo.Quantity ?? "1",
                        Unit = quoteInfo.Unit ?? "PCS",
                        QuotedPrice = quoteInfo.QuotedPrice ?? "0.00",
                        QuoteDate = DateTime.Now.ToString("yyyy-MM-dd"),
                        QuoteValidUntil = DateTime.Now.AddDays(30).ToString("yyyy-MM-dd"),
                        QuoteStatus = "Draft",
                        Priority = "Normal",
                        ExtractionMethod = "EMAIL_EXTRACTION",
                        ExtractionDomain = emailDomain
                    };

                    quotes.Add(quote);
                    processingLog.Add($"✅ Created quote for {articleCode}");
                }

                processingLog.Add($"📦 Total quotes created: {quotes.Count}");
                return quotes;
            }
            catch (Exception ex)
            {
                processingLog.Add($"❌ Error in ExtractQuotes: {ex.Message}");
                return new List<QuoteLine>();
            }
        }

        #endregion

        #region Private Helper Methods

        private static (string Description, string RfqNumber, string Quantity, string Unit, string QuotedPrice) ExtractQuoteDetailsInline(string subject, string emailBody)
        {
            var description = ExtractDescriptionInline(subject, emailBody);
            var rfqNumber = ExtractRfqNumberInline(subject, emailBody);
            var quantity = ExtractQuantityInline(subject, emailBody);
            var unit = ExtractUnitInline(emailBody);
            var quotedPrice = ExtractPriceInline(emailBody);

            return (description, rfqNumber, quantity, unit, quotedPrice);
        }

        private static string ExtractDescriptionInline(string subject, string emailBody)
        {
            if (!string.IsNullOrEmpty(subject) && subject.Length > 10)
                return subject.Substring(0, Math.Min(subject.Length, 100));

            return "Quote extracted from email";
        }

        private static string ExtractRfqNumberInline(string subject, string emailBody)
        {
            var patterns = new[]
            {
                @"RFQ[#:\s]*([A-Z0-9\-]{6,20})",
                @"Quote[#:\s]*([A-Z0-9\-]{6,20})",
                @"Request[#:\s]*([A-Z0-9\-]{6,20})"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(subject + " " + emailBody, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    return match.Groups[1].Value.Trim();
                }
            }

            return null;
        }

        private static string ExtractQuantityInline(string subject, string emailBody)
        {
            var patterns = new[]
            {
                @"qty[:\s]*(\d+)",
                @"quantity[:\s]*(\d+)",
                @"(\d+)\s*pcs",
                @"(\d+)\s*pieces"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(subject + " " + emailBody, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    return match.Groups[1].Value;
                }
            }

            return "1"; // Default quantity
        }

        private static string ExtractUnitInline(string emailBody)
        {
            var unitPatterns = new[] { "pcs", "pieces", "each", "ea", "st", "stuks" };

            foreach (var unit in unitPatterns)
            {
                if (emailBody.ToLower().Contains(unit))
                {
                    return unit.ToUpper();
                }
            }

            return "PCS";
        }

        private static string ExtractPriceInline(string emailBody)
        {
            var pricePatterns = new[]
            {
                @"price[:\s]*€?\$?(\d+[.,]\d{2})",
                @"€\s*(\d+[.,]\d{2})",
                @"\$\s*(\d+[.,]\d{2})"
            };

            foreach (var pattern in pricePatterns)
            {
                var match = Regex.Match(emailBody, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    return NormalizePrice(match.Groups[1].Value);
                }
            }

            return "0.00";
        }

        private static string NormalizePrice(string price)
        {
            if (string.IsNullOrEmpty(price)) return "0.00";
            return price.Replace(",", ".");
        }

        private static string GenerateRfqNumberInline(string subject, string articleCode)
        {
            var timestamp = DateTime.Now.ToString("yyyyMMddHHmm");
            var prefix = articleCode.Length >= 6 ? articleCode.Substring(0, 6).Replace(".", "") : "RFQ";
            return $"RFQ-{prefix}-{timestamp}";
        }

        #endregion

        #region Business Rule Validation

        /// <summary>
        /// Validate quote meets business requirements
        /// </summary>
        public static bool ValidateQuoteBusinessRules(QuoteLine quote)
        {
            var errors = new List<string>();

            // Business Rule 1: Article Code is mandatory
            if (string.IsNullOrWhiteSpace(quote.ArtiCode))
                errors.Add("Article Code is required");

            // Business Rule 2: RFQ Number is mandatory
            if (string.IsNullOrWhiteSpace(quote.RfqNumber))
                errors.Add("RFQ Number is required");

            // Business Rule 3: Valid quantity format
            if (!string.IsNullOrWhiteSpace(quote.Quantity))
            {
                if (!decimal.TryParse(quote.Quantity.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal qty) || qty <= 0)
                    errors.Add("Quantity must be a positive number");
            }

            // Business Rule 4: Valid pricing (if provided)
            if (!string.IsNullOrEmpty(quote.QuotedPrice))
            {
                if (!decimal.TryParse(quote.QuotedPrice.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal price) || price < 0)
                    errors.Add("Quoted price must be a valid positive number");
            }

            // Business Rule 5: Valid quote date
            if (!string.IsNullOrEmpty(quote.QuoteDate))
            {
                if (!DateTime.TryParse(quote.QuoteDate, out DateTime quoteDate))
                    errors.Add("Invalid quote date format");
            }

            if (errors.Any())
            {
                processingLog.Add($"Validation errors for {quote.ArtiCode}: {string.Join(", ", errors)}");
            }

            return errors.Count == 0;
        }

        /// <summary>
        /// Apply customer-specific quote business rules
        /// </summary>
        public static void ApplyQuoteBusinessRules(QuoteLine quote, string customerDomain)
        {
            // Business Rule: Set default values if missing
            if (string.IsNullOrEmpty(quote.Unit))
                quote.Unit = "PCS";

            if (string.IsNullOrEmpty(quote.QuoteStatus))
                quote.QuoteStatus = "Draft";

            if (string.IsNullOrEmpty(quote.Priority))
                quote.Priority = "Normal";

            // Business Rule: Set quote date if missing
            if (string.IsNullOrEmpty(quote.QuoteDate))
            {
                quote.QuoteDate = DateTime.Now.ToString("yyyy-MM-dd");
            }

            // Business Rule: Set quote validity period
            if (string.IsNullOrEmpty(quote.QuoteValidUntil))
            {
                quote.QuoteValidUntil = DateTime.Now.AddDays(30).ToString("yyyy-MM-dd"); // 30 days validity
            }

            // Business Rule: Customer-specific processing
            ApplyCustomerSpecificQuoteRules(quote, customerDomain);
        }

        /// <summary>
        /// Calculate quote pricing based on business rules
        /// </summary>
        public static void CalculateQuotePricing(QuoteLine quote)
        {
            try
            {
                if (decimal.TryParse(quote.Quantity.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal quantity) &&
                    decimal.TryParse(quote.QuotedPrice.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal unitPrice))
                {
                    var totalPrice = quantity * unitPrice;
                    // Note: If QuoteLine has TotalPrice property, set it here
                    // quote.TotalPrice = totalPrice.ToString("F2", CultureInfo.InvariantCulture);
                }
            }
            catch (Exception ex)
            {
                processingLog.Add($"Error calculating pricing for {quote.ArtiCode}: {ex.Message}");
            }
        }

        /// <summary>
        /// Assign priority based on business rules
        /// </summary>
        public static void AssignPriorityBasedOnBusinessRules(QuoteLine quote, string customerDomain)
        {
            // Business rules for priority assignment
            if (customerDomain?.Contains("weir") == true)
            {
                quote.Priority = "High"; // Weir gets high priority
            }
            else
            {
                quote.Priority = "Normal";
            }
        }

        /// <summary>
        /// Format quote for MKG standards
        /// </summary>
        public static void FormatQuoteForMkgStandards(QuoteLine quote)
        {
            // Ensure proper formatting for MKG system
            if (!string.IsNullOrEmpty(quote.QuotedPrice))
            {
                if (decimal.TryParse(quote.QuotedPrice, out decimal price))
                {
                    quote.QuotedPrice = price.ToString("F2", CultureInfo.InvariantCulture);
                }
            }
        }
        public static async Task<List<QuoteLine>> ExtractQuotesSafe(string emailBody, string emailDomain, string subject, object attachments)
        {
            try
            {
                await Task.Delay(1);
                return ExtractQuotes(subject, emailBody, emailDomain);
            }
            catch (Exception ex)
            {
                processingLog.Add($"Error in ExtractQuotesSafe: {ex.Message}");
                return new List<QuoteLine>();
            }
        }
        /// <summary>
        /// Apply customer-specific quote rules
        /// </summary>
        private static void ApplyCustomerSpecificQuoteRules(QuoteLine quote, string customerDomain)
        {
            if (customerDomain?.Contains("weir") == true)
            {
                // Weir-specific rules
                quote.QuoteValidUntil = DateTime.Now.AddDays(45).ToString("yyyy-MM-dd"); // Longer validity for Weir
            }
        }

        #endregion
    }
}