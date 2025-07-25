using System;
using System.Collections.Generic;
using System.Linq;
using Mkg_Elcotec_Automation.Models;
using Mkg_Elcotec_Automation.Models.EmailModels;

namespace Mkg_Elcotec_Automation.Services
{
    /// <summary>
    /// Enhanced Duplication Service - Complete Business Protection System
    /// 🛡️ Protects Elcotec from financial risks and data corruption
    /// 🎯 Smart categorization: Critical errors vs managed duplicates
    /// </summary>
    public static class DuplicationService
    {
        #region Events - Smart Categorization

        // 🚨 CRITICAL BUSINESS ERRORS - RED tab triggers
        public static event Action<int> OnCriticalBusinessErrors;

        // ❌ TECHNICAL INJECTION FAILURES - RED tab triggers  
        public static event Action<int> OnTechnicalInjectionFailures;

        // 🔄 MANAGED MKG DUPLICATES - YELLOW tab triggers
        public static event Action<int> OnManagedMkgDuplicates;

        // 💰 PRICE DISCREPANCY ALERTS - RED tab triggers
        public static event Action<PriceDiscrepancyAlert> OnPriceDiscrepancyDetected;

        // 📧 EMAIL-LEVEL DUPLICATES - Handled automatically
        public static event Action<int> OnEmailLevelDuplicatesDetected;

        // 🔄 CROSS-EMAIL DUPLICATES - Handled automatically  
        public static event Action<int> OnCrossEmailDuplicatesDetected;

        #endregion

        #region Tracking Collections - Business Intelligence

        // Financial protection tracking
        private static List<PriceDiscrepancyAlert> _priceDiscrepancies = new List<PriceDiscrepancyAlert>();
        private static List<CriticalBusinessError> _criticalBusinessErrors = new List<CriticalBusinessError>();

        // Managed duplicate tracking
        private static HashSet<string> _seenOrderKeys = new HashSet<string>();
        private static HashSet<string> _seenQuoteKeys = new HashSet<string>();
        private static HashSet<string> _seenRevisionKeys = new HashSet<string>();

        // Article price tracking for validation
        private static Dictionary<string, decimal> _knownArticlePrices = new Dictionary<string, decimal>();
        private static Dictionary<string, string> _articlePriceSources = new Dictionary<string, string>();

        #endregion

        #region Public Statistics - Business Dashboard

        public static int TotalCriticalBusinessErrors => _criticalBusinessErrors.Count;
        public static int TotalPriceDiscrepancies => _priceDiscrepancies.Count;
        public static int TotalManagedDuplicates { get; private set; } = 0;
        public static int TotalCrossEmailDuplicates { get; private set; } = 0;
        public static int TotalEmailLevelDuplicates { get; private set; } = 0;

        public static List<PriceDiscrepancyAlert> GetPriceDiscrepancies() => new List<PriceDiscrepancyAlert>(_priceDiscrepancies);
        public static List<CriticalBusinessError> GetCriticalBusinessErrors() => new List<CriticalBusinessError>(_criticalBusinessErrors);

        #endregion

        #region Reset and Initialization

        /// <summary>
        /// Reset all tracking for a new automation run
        /// 🔄 Clears all business protection tracking
        /// </summary>
        public static void ResetTrackingForNewRun()
        {
            // Clear duplicate tracking
            _seenOrderKeys.Clear();
            _seenQuoteKeys.Clear();
            _seenRevisionKeys.Clear();

            // Clear business protection tracking
            _priceDiscrepancies.Clear();
            _criticalBusinessErrors.Clear();
            _knownArticlePrices.Clear();
            _articlePriceSources.Clear();

            // Reset counters
            TotalManagedDuplicates = 0;
            TotalCrossEmailDuplicates = 0;
            TotalEmailLevelDuplicates = 0;

            Console.WriteLine("🛡️ DuplicationService: All business protection tracking reset for new run");
        }

        #endregion

        #region Email Processing - Smart Duplicate Detection

        /// <summary>
        /// Process email with enhanced business protection
        /// 🛡️ Detects all duplicate types and price discrepancies
        /// </summary>
        public static EmailDuplicationResult ProcessEmailWithBusinessProtection(EmailDetail emailDetail)
        {
            try
            {
                Console.WriteLine($"🔍 Enhanced Processing: '{emailDetail.Subject}' from {emailDetail.Sender}");

                var result = new EmailDuplicationResult
                {
                    OriginalEmail = emailDetail,
                    CleanedEmail = CreateCleanEmailCopy(emailDetail),
                    DuplicatesFound = new List<DuplicateItem>(),
                    PriceDiscrepancies = new List<PriceDiscrepancyAlert>(),
                    BusinessErrors = new List<CriticalBusinessError>()
                };

                // Step 1: HTML vs PDF Duplicate Detection (within same email)
                DetectHtmlVsPdfDuplicates(emailDetail, result);

                // Step 2: Cross-Email Duplicate Detection
                DetectCrossEmailDuplicates(emailDetail, result);

                // Step 3: Price Validation and Business Logic
                ValidateBusinessLogicAndPrices(emailDetail, result);

                // Step 4: Generate cleaned email (duplicates removed)
                GenerateCleanedEmail(emailDetail, result);

                // Step 5: Fire appropriate events based on findings
                FireBusinessProtectionEvents(result);

                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in enhanced email processing: {ex.Message}");

                // Create critical business error for processing failure
                var criticalError = new CriticalBusinessError
                {
                    ErrorType = "EMAIL_PROCESSING_FAILURE",
                    Description = $"Failed to process email '{emailDetail.Subject}': {ex.Message}",
                    FinancialRisk = "UNKNOWN",
                    RequiresManualReview = true,
                    EmailSource = emailDetail.Sender
                };

                _criticalBusinessErrors.Add(criticalError);
                OnCriticalBusinessErrors?.Invoke(1);

                throw;
            }
        }

        #endregion

        #region HTML vs PDF Duplicate Detection

        /// <summary>
        /// Detect HTML vs PDF duplicates within the same email
        /// 🔍 Compares extraction results from different sources
        /// </summary>
        private static void DetectHtmlVsPdfDuplicates(EmailDetail emailDetail, EmailDuplicationResult result)
        {
            try
            {
                if (emailDetail.Orders?.Count <= 1) return;

                // Group by ArtiCode to find potential HTML vs PDF duplicates
                var articleGroups = emailDetail.Orders
                    .Where(o => !string.IsNullOrEmpty(GetOrderProperty(o, "ArtiCode")))
                    .GroupBy(o => GetOrderProperty(o, "ArtiCode"))
                    .Where(g => g.Count() > 1)
                    .ToList();

                foreach (var group in articleGroups)
                {
                    var items = group.ToList();
                    var artiCode = group.Key;

                    // Check if we have different extraction methods (HTML vs PDF)
                    var extractionMethods = new List<string>();
                    foreach (var item in items)
                    {
                        var method = GetOrderProperty(item, "ExtractionMethod") ?? "UNKNOWN";
                        if (!extractionMethods.Contains(method))
                        {
                            extractionMethods.Add(method);
                        }
                    }

                    if (extractionMethods.Count > 1)
                    {
                        Console.WriteLine($"🔍 HTML vs PDF duplicate detected: {artiCode}");

                        // Validate prices between extraction methods
                        ValidateHtmlVsPdfPrices(artiCode, items, result);

                        // Create duplicate entry
                        var duplicate = new DuplicateItem
                        {
                            DuplicateType = "HTML_VS_PDF",
                            ArtiCode = artiCode,
                            Description = $"Same article extracted from multiple sources: {string.Join(", ", extractionMethods)}",
                            Sources = extractionMethods, // Now this is List<string>
                            ItemCount = items.Count,
                            EmailSource = emailDetail.Sender,
                            RequiresPriceValidation = true
                        };

                        result.DuplicatesFound.Add(duplicate);
                        TotalEmailLevelDuplicates++;
                    }
                }

                // Fire event if HTML vs PDF duplicates were found
                var htmlPdfCount = 0;
                foreach (var dup in result.DuplicatesFound)
                {
                    if (dup.DuplicateType == "HTML_VS_PDF")
                    {
                        htmlPdfCount++;
                    }
                }

                if (htmlPdfCount > 0)
                {
                    OnEmailLevelDuplicatesDetected?.Invoke(htmlPdfCount);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error detecting HTML vs PDF duplicates: {ex.Message}");
            }
        }

        /// <summary>
        /// Validate prices between HTML and PDF extraction
        /// 💰 Critical business protection against pricing errors
        /// </summary>
        private static void ValidateHtmlVsPdfPrices(string artiCode, List<dynamic> items, EmailDuplicationResult result)
        {
            try
            {
                var priceVariations = new List<(string source, decimal price)>();

                foreach (var item in items)
                {
                    var priceStr = GetOrderProperty(item, "UnitPrice") ?? GetOrderProperty(item, "TotalPrice") ?? "0";
                    var source = GetOrderProperty(item, "ExtractionMethod") ?? "UNKNOWN";

                    if (decimal.TryParse(priceStr, out decimal price) && price > 0)
                    {
                        priceVariations.Add((source, price));
                    }
                }

                if (priceVariations.Count > 1)
                {
                    var minPrice = priceVariations.Min(p => p.price);
                    var maxPrice = priceVariations.Max(p => p.price);
                    var priceDiscrepancy = Math.Abs(maxPrice - minPrice);
                    var discrepancyPercentage = minPrice > 0 ? (priceDiscrepancy / minPrice) * 100 : 0;

                    // 🚨 CRITICAL: Price discrepancy > 5% requires manual review
                    if (discrepancyPercentage > 5)
                    {
                        var alert = new PriceDiscrepancyAlert
                        {
                            ArtiCode = artiCode,
                            PriceRange = $"€{minPrice:F2} - €{maxPrice:F2}",
                            DiscrepancyAmount = priceDiscrepancy,
                            DiscrepancyPercentage = discrepancyPercentage,
                            Sources = priceVariations.Select(p => $"{p.source}: €{p.price:F2}").ToList(),
                            FinancialRisk = discrepancyPercentage > 20 ? "HIGH" : "MEDIUM",
                            RequiresManualReview = true,
                            DetectionMethod = "HTML_VS_PDF_COMPARISON"
                        };

                        result.PriceDiscrepancies.Add(alert);
                        _priceDiscrepancies.Add(alert);

                        Console.WriteLine($"🚨 PRICE DISCREPANCY: {artiCode} - {discrepancyPercentage:F1}% difference");
                        OnPriceDiscrepancyDetected?.Invoke(alert);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error validating HTML vs PDF prices: {ex.Message}");
            }
        }

        #endregion

        #region Cross-Email Duplicate Detection

        /// <summary>
        /// Detect duplicates across different emails
        /// 🔄 Prevents processing the same order/quote/revision multiple times
        /// </summary>
        private static void DetectCrossEmailDuplicates(EmailDetail emailDetail, EmailDuplicationResult result)
        {
            try
            {
                if (emailDetail.Orders?.Any() != true) return;

                foreach (var order in emailDetail.Orders)
                {
                    var artiCode = GetOrderProperty(order, "ArtiCode");
                    var poNumber = GetOrderProperty(order, "PoNumber");
                    var rfqNumber = GetOrderProperty(order, "RfqNumber");

                    if (string.IsNullOrEmpty(artiCode)) continue;

                    // Create unique keys for different business types
                    string orderKey = null;
                    string quoteKey = null;
                    string revisionKey = null;
                    string duplicateType = null;

                    // Determine the type and create appropriate key
                    if (!string.IsNullOrEmpty(poNumber))
                    {
                        orderKey = $"ORDER|{poNumber}|{artiCode}";
                        duplicateType = "CROSS_EMAIL_ORDER";
                    }
                    else if (!string.IsNullOrEmpty(rfqNumber))
                    {
                        quoteKey = $"QUOTE|{rfqNumber}|{artiCode}";
                        duplicateType = "CROSS_EMAIL_QUOTE";
                    }
                    else if (HasRevisionData(order))
                    {
                        var currentRev = GetOrderProperty(order, "CurrentRevision");
                        var newRev = GetOrderProperty(order, "NewRevision");
                        revisionKey = $"REVISION|{artiCode}|{currentRev}|{newRev}";
                        duplicateType = "CROSS_EMAIL_REVISION";
                    }

                    // Check for cross-email duplicates
                    bool isDuplicate = false;

                    if (orderKey != null && _seenOrderKeys.Contains(orderKey))
                    {
                        isDuplicate = true;
                    }
                    else if (quoteKey != null && _seenQuoteKeys.Contains(quoteKey))
                    {
                        isDuplicate = true;
                    }
                    else if (revisionKey != null && _seenRevisionKeys.Contains(revisionKey))
                    {
                        isDuplicate = true;
                    }

                    if (isDuplicate)
                    {
                        var duplicate = new DuplicateItem
                        {
                            DuplicateType = duplicateType,
                            ArtiCode = artiCode,
                            Description = $"Already processed in previous email",
                            EmailSource = emailDetail.Sender,
                            ItemCount = 1,
                            RequiresPriceValidation = false,
                            HandledAutomatically = true
                        };

                        result.DuplicatesFound.Add(duplicate);
                        TotalCrossEmailDuplicates++;

                        Console.WriteLine($"🔄 Cross-email duplicate: {duplicateType} - {artiCode}");
                    }
                    else
                    {
                        // Add to tracking sets
                        if (orderKey != null) _seenOrderKeys.Add(orderKey);
                        if (quoteKey != null) _seenQuoteKeys.Add(quoteKey);
                        if (revisionKey != null) _seenRevisionKeys.Add(revisionKey);
                    }
                }

                var crossEmailDuplicateCount = result.DuplicatesFound.Count(d => d.DuplicateType.StartsWith("CROSS_EMAIL"));
                if (crossEmailDuplicateCount > 0)
                {
                    OnCrossEmailDuplicatesDetected?.Invoke(crossEmailDuplicateCount);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error detecting cross-email duplicates: {ex.Message}");
            }
        }

        #endregion

        #region Business Logic Validation

        /// <summary>
        /// Validate business logic and detect critical errors
        /// 🛡️ Financial protection and business rule enforcement
        /// </summary>
        private static void ValidateBusinessLogicAndPrices(EmailDetail emailDetail, EmailDuplicationResult result)
        {
            try
            {
                if (emailDetail.Orders?.Any() != true) return;

                foreach (var order in emailDetail.Orders)
                {
                    var artiCode = GetOrderProperty(order, "ArtiCode");
                    if (string.IsNullOrEmpty(artiCode)) continue;

                    // Validate article price consistency
                    ValidateArticlePriceConsistency(order, result);

                    // Validate business rules
                    ValidateBusinessRules(order, emailDetail, result);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in business logic validation: {ex.Message}");
            }
        }

        /// <summary>
        /// Validate article price consistency across the system
        /// 💰 Prevents financial errors and pricing inconsistencies
        /// </summary>
        private static void ValidateArticlePriceConsistency(dynamic order, EmailDuplicationResult result)
        {
            try
            {
                var artiCode = GetOrderProperty(order, "ArtiCode");
                var priceStr = GetOrderProperty(order, "UnitPrice") ?? "0";
                var source = GetOrderProperty(order, "ExtractionMethod") ?? "UNKNOWN";

                if (!decimal.TryParse(priceStr, out decimal currentPrice) || currentPrice <= 0)
                    return;

                if (_knownArticlePrices.ContainsKey(artiCode))
                {
                    var knownPrice = _knownArticlePrices[artiCode];
                    var knownSource = _articlePriceSources.ContainsKey(artiCode) ? _articlePriceSources[artiCode] : "UNKNOWN";
                    var priceDifference = Math.Abs(currentPrice - knownPrice);
                    var percentageDifference = knownPrice > 0 ? (priceDifference / knownPrice) * 100 : 0;

                    // 🚨 CRITICAL: Price discrepancy > 10% for known articles
                    if (percentageDifference > 10)
                    {
                        var alert = new PriceDiscrepancyAlert
                        {
                            ArtiCode = artiCode,
                            PriceRange = $"Known: €{knownPrice:F2}, Current: €{currentPrice:F2}",
                            DiscrepancyAmount = priceDifference,
                            DiscrepancyPercentage = percentageDifference,
                            Sources = new List<string> { $"Known ({knownSource}): €{knownPrice:F2}", $"Current ({source}): €{currentPrice:F2}" },
                            FinancialRisk = percentageDifference > 25 ? "HIGH" : "MEDIUM",
                            RequiresManualReview = true,
                            DetectionMethod = "PRICE_CONSISTENCY_CHECK"
                        };

                        result.PriceDiscrepancies.Add(alert);
                        _priceDiscrepancies.Add(alert);
                        OnPriceDiscrepancyDetected?.Invoke(alert);
                    }
                }
                else
                {
                    // Record this price for future validation
                    _knownArticlePrices[artiCode] = currentPrice;
                    _articlePriceSources[artiCode] = source;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error validating price consistency: {ex.Message}");
            }
        }

        /// <summary>
        /// Validate critical business rules
        /// 🛡️ Protects against business logic violations
        /// </summary>
        private static void ValidateBusinessRules(dynamic order, EmailDetail emailDetail, EmailDuplicationResult result, EnhancedProgressManager progressManager = null)
        {
            try
            {
                var artiCode = GetOrderProperty(order, "ArtiCode");
                var quantity = GetOrderProperty(order, "Quantity");
                var price = GetOrderProperty(order, "UnitPrice");

                // Rule 1: Critical articles must have valid quantities
                if (IsHighValueArticle(artiCode))
                {
                    if (!int.TryParse(quantity, out int qty) || qty <= 0)
                    {
                        var error = new CriticalBusinessError
                        {
                            ErrorType = "INVALID_QUANTITY_HIGH_VALUE_ARTICLE",
                            Description = $"High-value article {artiCode} has invalid quantity: {quantity}",
                            FinancialRisk = "HIGH",
                            RequiresManualReview = true,
                            EmailSource = emailDetail.Sender,
                            ArtiCode = artiCode
                        };

                        result.BusinessErrors.Add(error);
                        _criticalBusinessErrors.Add(error);
                        // 🎯 INCREMENT B.ERRORS for critical business rule violation
                        progressManager?.IncrementBusinessErrors("Pre-Validation");
                    }
                }

                // Rule 2: Validate price ranges for known article categories
                if (decimal.TryParse(price, out decimal unitPrice))
                {
                    var priceValidation = ValidateArticlePriceRange(artiCode, unitPrice);
                    if (!priceValidation.IsValid)
                    {
                        var error = new CriticalBusinessError
                        {
                            ErrorType = "PRICE_OUTSIDE_EXPECTED_RANGE",
                            Description = priceValidation.ErrorMessage,
                            FinancialRisk = unitPrice > 10000 ? "HIGH" : "MEDIUM",
                            RequiresManualReview = true,
                            EmailSource = emailDetail.Sender,
                            ArtiCode = artiCode
                        };

                        result.BusinessErrors.Add(error);
                        _criticalBusinessErrors.Add(error);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error validating business rules: {ex.Message}");
            }
        }

        #endregion

        #region Helper Methods

        private static string GetOrderProperty(dynamic order, string propertyName)
        {
            try
            {
                var type = order.GetType();
                var property = type.GetProperty(propertyName);
                return property?.GetValue(order)?.ToString();
            }
            catch
            {
                return null;
            }
        }

        private static bool HasRevisionData(dynamic order)
        {
            var currentRev = GetOrderProperty(order, "CurrentRevision");
            var newRev = GetOrderProperty(order, "NewRevision");
            return !string.IsNullOrEmpty(currentRev) && !string.IsNullOrEmpty(newRev);
        }

        private static EmailDetail CreateCleanEmailCopy(EmailDetail original)
        {
            return new EmailDetail
            {
                Subject = original.Subject,
                Sender = original.Sender,
                ReceivedDate = original.ReceivedDate,
                ClientDomain = original.ClientDomain,
                Orders = new List<dynamic>()
            };
        }

        private static void GenerateCleanedEmail(EmailDetail originalEmail, EmailDuplicationResult result)
        {
            try
            {
                var duplicateArticles = result.DuplicatesFound.Select(d => d.ArtiCode).ToHashSet();

                if (originalEmail.Orders?.Any() == true)
                {
                    foreach (var order in originalEmail.Orders)
                    {
                        var artiCode = GetOrderProperty(order, "ArtiCode");

                        // Only include non-duplicate items in cleaned email
                        if (!duplicateArticles.Contains(artiCode))
                        {
                            result.CleanedEmail.Orders.Add(order);
                        }
                    }
                }

                Console.WriteLine($"📧 Cleaned email: {originalEmail.Orders?.Count ?? 0} → {result.CleanedEmail.Orders.Count} items");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error generating cleaned email: {ex.Message}");
                // Fall back to original email
                result.CleanedEmail = originalEmail;
            }
        }

        private static void FireBusinessProtectionEvents(EmailDuplicationResult result)
        {
            try
            {
                // Critical business errors
                if (result.BusinessErrors?.Any() == true)
                {
                    OnCriticalBusinessErrors?.Invoke(result.BusinessErrors.Count);
                }

                // Managed duplicates (automatically handled)
                var managedDuplicates = result.DuplicatesFound?.Count(d => d.HandledAutomatically) ?? 0;
                if (managedDuplicates > 0)
                {
                    TotalManagedDuplicates += managedDuplicates;
                    OnManagedMkgDuplicates?.Invoke(managedDuplicates);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error firing business protection events: {ex.Message}");
            }
        }

        private static bool IsHighValueArticle(string artiCode)
        {
            // Define high-value article patterns
            if (string.IsNullOrEmpty(artiCode)) return false;

            return artiCode.StartsWith("HV-") ||
                   artiCode.Contains("PUMP") ||
                   artiCode.Contains("VALVE") ||
                   artiCode.Contains("MOTOR");
        }

        private static (bool IsValid, string ErrorMessage) ValidateArticlePriceRange(string artiCode, decimal price)
        {
            try
            {
                // Basic price validation rules
                if (price <= 0)
                    return (false, $"Article {artiCode} has invalid price: €{price}");

                if (price > 50000)
                    return (false, $"Article {artiCode} price €{price:F2} exceeds maximum threshold (€50,000)");

                // Category-specific validation could be added here
                return (true, "");
            }
            catch
            {
                return (false, $"Unable to validate price for {artiCode}");
            }
        }

        #endregion
    }

    #region Supporting Models

    public class EmailDuplicationResult
    {
        public EmailDetail OriginalEmail { get; set; }
        public EmailDetail CleanedEmail { get; set; }
        public List<DuplicateItem> DuplicatesFound { get; set; } = new List<DuplicateItem>();
        public List<PriceDiscrepancyAlert> PriceDiscrepancies { get; set; } = new List<PriceDiscrepancyAlert>();
        public List<CriticalBusinessError> BusinessErrors { get; set; } = new List<CriticalBusinessError>();

        public bool HasCriticalIssues => (BusinessErrors?.Any() == true) ||
                                        (PriceDiscrepancies?.Any(p => p.FinancialRisk == "HIGH") == true);

        public bool HasManagedDuplicates => DuplicatesFound?.Any(d => d.HandledAutomatically) == true;

        public string GetSummary()
        {
            var critical = BusinessErrors?.Count ?? 0;
            var priceIssues = PriceDiscrepancies?.Count ?? 0;
            var duplicates = DuplicatesFound?.Count ?? 0;

            return $"Critical: {critical}, Price Issues: {priceIssues}, Duplicates: {duplicates}";
        }
    }

    public class DuplicateItem
    {
        public string DuplicateType { get; set; } // HTML_VS_PDF, CROSS_EMAIL_ORDER, etc.
        public string ArtiCode { get; set; }
        public string Description { get; set; }
        public List<string> Sources { get; set; } = new List<string>();
        public int ItemCount { get; set; }
        public string EmailSource { get; set; }
        public bool RequiresPriceValidation { get; set; }
        public bool HandledAutomatically { get; set; }
    }

    public class PriceDiscrepancyAlert
    {
        public string ArtiCode { get; set; }
        public string PriceRange { get; set; }
        public decimal DiscrepancyAmount { get; set; }
        public decimal DiscrepancyPercentage { get; set; }
        public List<string> Sources { get; set; } = new List<string>();
        public string FinancialRisk { get; set; } // HIGH, MEDIUM, LOW
        public bool RequiresManualReview { get; set; }
        public string DetectionMethod { get; set; }
        public DateTime DetectedAt { get; set; } = DateTime.Now;
    }

    public class CriticalBusinessError
    {
        public string ErrorType { get; set; }
        public string Description { get; set; }
        public string FinancialRisk { get; set; } // HIGH, MEDIUM, LOW
        public bool RequiresManualReview { get; set; }
        public string EmailSource { get; set; }
        public string ArtiCode { get; set; }
        public DateTime DetectedAt { get; set; } = DateTime.Now;
    }

    #endregion
}