using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Graph.Models;
using Mkg_Elcotec_Automation.Models;

namespace Mkg_Elcotec_Automation.Services
{
    /// <summary>
    /// REFACTORED: QuoteLogicService - Focuses purely on quote business logic
    /// Domain checking and email content analysis removed - handled by EmailContentAnalyzer and EnhancedDomainProcessor
    /// This service now handles: validation, processing, calculations, formatting, and business rules
    /// </summary>
    public static class QuoteLogicService
    {
        private static List<string> processingLog = new List<string>();

        #region Public Properties
        public static List<string> GetProcessingLog() => new List<string>(processingLog);
        public static void ClearProcessingLog() => processingLog.Clear();
        #endregion

        #region Core Quote Business Logic
        #region Missing Method Implementations

        /// <summary>
        /// Check if email contains quote/RFQ content
        /// </summary>
        public static bool IsQuoteEmail(string subject, string emailBody, AttachmentCollectionResponse attachments)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(subject))
                    return false;

                var subjectLower = subject.ToLower();
                var bodyLower = emailBody?.ToLower() ?? "";

                // SIMPLIFIED: Just check if subject contains basic quote keywords
                var quoteKeywords = new[]{
                                    "rfq",              // Request for Quote
                                    "request for quote",
                                    "offerte",          // Dutch: quote
                                    "quotation",
                                    "quote",
                                    "pricing",
                                    "price inquiry",
                                    "kostprijs"         // Dutch: cost price
                                };

                // Check for quote keywords in subject (most important)
                bool hasQuoteKeyword = quoteKeywords.Any(keyword => subjectLower.Contains(keyword));

                if (hasQuoteKeyword)
                {
                    Console.WriteLine($"✅ IsQuoteEmail = TRUE for: {subject} (found keyword)");
                    return true;
                }

                Console.WriteLine($"❌ IsQuoteEmail = FALSE for: {subject}");
                return false;
            }
            catch (Exception ex)
            {
                processingLog.Add($"Error in IsQuoteEmail: {ex.Message}");
                Console.WriteLine($"❌ IsQuoteEmail ERROR for: {subject} - {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Extract quotes from email content
        /// </summary>
        public static List<QuoteLine> ExtractQuotes(string emailBody, string emailDomain, string subject, AttachmentCollectionResponse attachments)
        {
            try
            {
                processingLog.Add("=== EXTRACTING QUOTES FROM EMAIL ===");
                Console.WriteLine($"🔍 ExtractQuotes called for: {subject}");

                var quotes = new List<QuoteLine>();

                // Check if this is a quote email
                if (!IsQuoteEmail(subject, emailBody, attachments))
                {
                    processingLog.Add("Not a quote email - returning empty list");
                    Console.WriteLine($"❌ Skipping {subject} - not recognized as quote email");
                    return quotes;
                }

                Console.WriteLine($"✅ Processing {subject} as quote email");

                // Extract article codes from subject
                var articleCodes = RevisionLogicService.ExtractArticleCodesFromSubject(subject);
                processingLog.Add($"Found {articleCodes.Count()} article codes for quote processing");
                Console.WriteLine($"   Article codes found: {string.Join(", ", articleCodes)}");

                // If no article codes found, create generic quote entry
                if (articleCodes.Count() == 0)
                {
                    Console.WriteLine("   No article codes found, creating generic quote entry");
                    articleCodes.Add("QUOTE-ITEM-" + DateTime.Now.ToString("HHmmss"));
                }

                // Extract quote information from subject/body
                var quoteInfo = ExtractQuoteDetails(subject, emailBody);

                // Create quotes for each article code found
                foreach (var articleCode in articleCodes)
                {
                    var quote = new QuoteLine
                    {
                        ArtiCode = articleCode,
                        Description = quoteInfo.Description ?? $"Quote for {articleCode}",
                        RfqNumber = quoteInfo.RfqNumber ?? GenerateRfqNumber(subject, articleCode),
                        Quantity = quoteInfo.Quantity ?? "1",
                        Unit = quoteInfo.Unit ?? "ST",
                        QuotedPrice = "0.00",
                        QuoteDate = DateTime.Now.ToString("yyyy-MM-dd"),
                        RequestedDeliveryDate = quoteInfo.RequestedDelivery ?? "",
                        QuoteStatus = "Draft",
                        Priority = DetermineQuotePriority(subject, emailBody)
                    };

                    quote.SetExtractionDetails("EMAIL_EXTRACTION", emailDomain);
                    quotes.Add(quote);

                    processingLog.Add($"✓ Created quote: {articleCode} (RFQ: {quote.RfqNumber})");
                    Console.WriteLine($"   ✅ Created quote: {articleCode} (RFQ: {quote.RfqNumber})");
                }

                processingLog.Add($"Extracted {quotes.Count} quotes from email");
                Console.WriteLine($"🎉 ExtractQuotes completed: {quotes.Count} quotes created");
                return quotes;
            }
            catch (Exception ex)
            {
                processingLog.Add($"Error in ExtractQuotes: {ex.Message}");
                Console.WriteLine($"❌ ExtractQuotes ERROR: {ex.Message}");
                return new List<QuoteLine>();
            }
        }

        /// <summary>
        /// Extract article codes from email subject
        /// </summary>
        public static (string RfqNumber, string Description, string Quantity, string Unit, string RequestedDelivery) ExtractQuoteDetails(string subject, string body)
        {
            // Extract RFQ number
            string rfqNumber = null;
            var rfqPattern = @"rfq[:\-\s#]*(\d+)";
            var rfqMatch = Regex.Match(subject, rfqPattern, RegexOptions.IgnoreCase);
            if (rfqMatch.Success)
            {
                rfqNumber = "RFQ-" + rfqMatch.Groups[1].Value;
            }

            // Extract quantity
            var qtyPattern = @"(qty|quantity|aantal)[:\-\s]*(\d+)";
            var qtyMatch = Regex.Match(subject + " " + body, qtyPattern, RegexOptions.IgnoreCase);
            string quantity = qtyMatch.Success ? qtyMatch.Groups[2].Value : "1";

            // Extract unit
            var unitPattern = @"(\d+)\s*(st|pcs|pieces|stuks|each)";
            var unitMatch = Regex.Match(body, unitPattern, RegexOptions.IgnoreCase);
            string unit = unitMatch.Success ? unitMatch.Groups[2].Value.ToUpper() : "ST";

            // Extract description
            string description = null;
            if (body.Contains("description:") || body.Contains("omschrijving:"))
            {
                var descPattern = @"(description|omschrijving):\s*([^\n\r]+)";
                var descMatch = Regex.Match(body, descPattern, RegexOptions.IgnoreCase);
                if (descMatch.Success)
                {
                    description = descMatch.Groups[2].Value.Trim();
                }
            }

            // Extract requested delivery
            string requestedDelivery = null;
            if (body.Contains("delivery") || body.Contains("levering"))
            {
                var deliveryPattern = @"(delivery|levering)[^:]*:\s*([^\n\r]+)";
                var deliveryMatch = Regex.Match(body, deliveryPattern, RegexOptions.IgnoreCase);
                if (deliveryMatch.Success)
                {
                    requestedDelivery = deliveryMatch.Groups[2].Value.Trim();
                }
            }

            return (rfqNumber, description, quantity, unit, requestedDelivery);
        }


        /// <summary>
        /// Generate RFQ number if not found in email
        /// </summary>
        public static string GenerateRfqNumber(string subject, string articleCode)
        {
            var timestamp = DateTime.Now.ToString("yyyyMMdd");
            var cleanSubject = Regex.Replace(subject, @"[^a-zA-Z0-9]", "").ToUpper();
            if (cleanSubject.Length > 10) cleanSubject = cleanSubject.Substring(0, 10);

            return $"RFQ-{timestamp}-{cleanSubject}-{articleCode.Replace(".", "")}";
        }

        /// <summary>
        /// Determine quote priority based on content
        /// </summary>
        public static string DetermineQuotePriority(string subject, string body)
        {
            var contentLower = (subject + " " + body).ToLower();

            // High priority indicators
            if (contentLower.Contains("urgent") || contentLower.Contains("immediate") ||
                contentLower.Contains("rush") || contentLower.Contains("asap") ||
                contentLower.Contains("spoedig") || contentLower.Contains("dringend"))
            {
                return "High";
            }

            // Low priority indicators
            if (contentLower.Contains("future") || contentLower.Contains("planning") ||
                contentLower.Contains("toekomst") || contentLower.Contains("planning"))
            {
                return "Low";
            }

            return "Normal";
        }


        #endregion
        /// <summary>
        /// Validate quote meets business requirements
        /// </summary>
        public static bool ValidateQuoteBusinessRules(QuoteLine quote)
        {
            var errors = new List<string>();

            // Business Rule 1: Article Code is mandatory
            if (string.IsNullOrWhiteSpace(quote.ArtiCode))
                errors.Add("Article Code is required");

            // Business Rule 2: RFQ Number is mandatory for quotes
            if (string.IsNullOrWhiteSpace(quote.RfqNumber))
                errors.Add("RFQ Number is required");

            // Business Rule 3: Valid quantity format and value
            if (string.IsNullOrWhiteSpace(quote.Quantity))
            {
                errors.Add("Quantity is required");
            }
            else
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
                if (decimal.TryParse(quote.Quantity.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal qty))
                {
                    // Business Rule: Apply volume pricing
                    if (qty >= 100)
                    {
                        processingLog.Add($"Volume pricing applicable for {quote.ArtiCode} (Qty: {qty})");
                    }

                    // Business Rule: Minimum quote value
                    if (!string.IsNullOrEmpty(quote.QuotedPrice))
                    {
                        if (decimal.TryParse(quote.QuotedPrice.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal price))
                        {
                            var totalValue = price * qty;
                            // Remove this line since TotalQuoteValue doesn't exist:
                            // quote.TotalQuoteValue = totalValue.ToString("F2");

                            // Business Rule: Minimum quote value check
                            if (totalValue < 50.00m)
                            {
                                processingLog.Add($"Quote {quote.ArtiCode} below minimum value (€50.00)");
                            }

                            // Log the total value instead
                            processingLog.Add($"Calculated total quote value for {quote.ArtiCode}: €{totalValue:F2}");
                        }
                    }
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
            try
            {
                // Business Rule: High-priority customers
                if (!string.IsNullOrEmpty(customerDomain))
                {
                    var domain = customerDomain.ToLower();
                    if (domain.Contains("weir") || domain.Contains("shell") || domain.Contains("bp"))
                    {
                        quote.Priority = "High";
                        processingLog.Add($"Applied high-priority customer rule to quote {quote.ArtiCode}");
                    }
                }

                // Business Rule: Urgent quotes
                if (quote.Description?.ToLower().Contains("urgent") == true ||
                    quote.Description?.ToLower().Contains("emergency") == true)
                {
                    quote.Priority = "High";
                    processingLog.Add($"Urgent quote detected for {quote.ArtiCode}");
                }

                // Business Rule: Large quantity quotes
                if (decimal.TryParse(quote.Quantity.Replace(",", "."), out decimal qty) && qty >= 50)
                {
                    quote.Priority = "High";
                    processingLog.Add($"Large quantity quote detected: {quote.ArtiCode} (Qty: {qty})");
                }
            }
            catch (Exception ex)
            {
                processingLog.Add($"Error assigning priority for {quote.ArtiCode}: {ex.Message}");
                quote.Priority = "Normal"; // Fallback
            }
        }

        /// <summary>
        /// Format quote data according to MKG standards
        /// </summary>
        public static void FormatQuoteForMkgStandards(QuoteLine quote)
        {
            // MKG Standard: Article code format
            if (!string.IsNullOrEmpty(quote.ArtiCode))
            {
                quote.ArtiCode = quote.ArtiCode.Trim().ToUpper();
            }

            // MKG Standard: RFQ number format
            if (!string.IsNullOrEmpty(quote.RfqNumber))
            {
                quote.RfqNumber = quote.RfqNumber.Trim().ToUpper();
            }

            // MKG Standard: Unit standardization
            quote.Unit = StandardizeUnit(quote.Unit);

            // MKG Standard: Date formats (yyyy-MM-dd)
            if (!string.IsNullOrEmpty(quote.QuoteDate))
            {
                if (DateTime.TryParse(quote.QuoteDate, out DateTime quoteDate))
                {
                    quote.QuoteDate = quoteDate.ToString("yyyy-MM-dd");
                }
            }

            if (!string.IsNullOrEmpty(quote.QuoteValidUntil))
            {
                if (DateTime.TryParse(quote.QuoteValidUntil, out DateTime validDate))
                {
                    quote.QuoteValidUntil = validDate.ToString("yyyy-MM-dd");
                }
            }

            // MKG Standard: Currency formatting
            if (!string.IsNullOrEmpty(quote.QuotedPrice))
            {
                if (decimal.TryParse(quote.QuotedPrice.Replace(",", "."), out decimal price))
                {
                    quote.QuotedPrice = price.ToString("F2");
                }
            }

            // MKG Standard: Description cleanup
            if (!string.IsNullOrEmpty(quote.Description))
            {
                quote.Description = CleanDescription(quote.Description);
            }
        }

        /// <summary>
        /// Apply customer-specific quote rules
        /// </summary>
        private static void ApplyCustomerSpecificQuoteRules(QuoteLine quote, string customerDomain)
        {
            if (string.IsNullOrEmpty(customerDomain)) return;

            var domain = customerDomain.ToLower();

            // Weir-specific rules
            if (domain.Contains("weir"))
            {
                quote.Priority = "High";
                quote.QuoteValidUntil = DateTime.Now.AddDays(60).ToString("yyyy-MM-dd"); // Extended validity
                processingLog.Add($"Applied Weir-specific quote rules to {quote.ArtiCode}");
            }

            // Shell-specific rules
            if (domain.Contains("shell"))
            {
                quote.Priority = "High";
                quote.QuoteValidUntil = DateTime.Now.AddDays(45).ToString("yyyy-MM-dd");
            }
        }

        /// <summary>
        /// Standardize unit format
        /// </summary>
        private static string StandardizeUnit(string unit)
        {
            if (string.IsNullOrEmpty(unit)) return "PCS";

            var standardUnits = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "EA", "PCS" },
        { "EACH", "PCS" },
        { "PIECE", "PCS" },
        { "STUKS", "PCS" },
        { "ST", "PCS" },
        { "STK", "PCS" }
    };

            return standardUnits.TryGetValue(unit.Trim(), out string standardUnit) ? standardUnit : unit.ToUpper();
        }

        /// <summary>
        /// Clean description text
        /// </summary>
        private static string CleanDescription(string description)
        {
            if (string.IsNullOrEmpty(description)) return "";

            description = Regex.Replace(description, @"\s+", " ").Trim();
            description = Regex.Replace(description, @"<[^>]*>", "");

            if (description.Length > 200)
                description = description.Substring(0, 197) + "...";

            return description;
        }
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

        // Add the rest of the methods following the same pattern as RevisionLogicService and OrderLogicService...
        // [Continue with ValidateQuoteBusinessRules, ApplyQuoteBusinessRules, etc.]

        #endregion

        #region Test Data Generation

        /// <summary>
        /// Generate test quotes for development/testing
        /// </summary>
        public static List<QuoteLine> GenerateTestQuotes()
        {
            var testQuotes = new List<QuoteLine>();

            try
            {
                // High-priority quote
                var urgentQuote = new QuoteLine();
                urgentQuote.ArtiCode = "345.678.901";
                urgentQuote.Description = "Emergency valve quote for offshore";
                urgentQuote.RfqNumber = "RFQ-2025-001";
                urgentQuote.Quantity = "5";
                urgentQuote.Unit = "PCS";
                urgentQuote.QuotedPrice = "2500.00";
                urgentQuote.Priority = "High";
                urgentQuote.QuoteStatus = "Draft";
                urgentQuote.SetExtractionDetails("TEST_GENERATION", "test.domain.com");

                testQuotes.Add(urgentQuote);

                processingLog.Add($"Generated {testQuotes.Count} test quotes for development");
                return testQuotes;
            }
            catch (Exception ex)
            {
                processingLog.Add($"Error generating test quotes: {ex.Message}");
                return new List<QuoteLine>();
            }
        }

        #endregion

        public static async Task<List<QuoteLine>> ExtractQuotesSafe(string emailBody, string emailDomain, string subject, AttachmentCollectionResponse attachments)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(emailBody) || string.IsNullOrWhiteSpace(emailDomain))
                {
                    return new List<QuoteLine>();
                }

                if (!emailDomain.StartsWith("@"))
                {
                    emailDomain = "@" + emailDomain;
                }

                // Use the existing IsQuoteEmail method with correct signature
                if (!QuoteLogicService.IsQuoteEmail(subject, emailBody, attachments))
                {
                    return new List<QuoteLine>();
                }

                // Use the existing ExtractQuotes method
                var quotes = QuoteLogicService.ExtractQuotes(emailBody, emailDomain, subject, attachments);

                return quotes;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in ExtractQuotesSafe: {ex.Message}");
                return new List<QuoteLine>();
            }
        }
    }
}