using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Graph.Models;
using Mkg_Elcotec_Automation.Models;

namespace Mkg_Elcotec_Automation.Services
{
    /// <summary>
    /// Enhanced Domain-Specific Email Processing with improved customer mapping and detection
    /// </summary>
    public static class EnhancedDomainProcessor
    {
        #region Domain Configuration & Mapping

        public class DomainConfig
        {
            public string Domain { get; set; }
            public string CustomerName { get; set; }
            public string AdministrationNumber { get; set; }
            public string DebtorNumber { get; set; }
            public string RelationNumber { get; set; }
            public List<string> EmailKeywords { get; set; } = new List<string>();
            public List<string> OrderKeywords { get; set; } = new List<string>();
            public List<string> QuoteKeywords { get; set; } = new List<string>();
            public List<string> RevisionKeywords { get; set; } = new List<string>();
            public bool IsHighPriority { get; set; } = false;
            public string SpecialParsingMethod { get; set; }
        }

        // Enhanced domain mappings with detailed configurations
        private static readonly Dictionary<string, DomainConfig> _domainMappings = new Dictionary<string, DomainConfig>
        {
            // Weir Group (Primary Customer)
            ["mail.weir"] = new DomainConfig
            {
                Domain = "mail.weir",
                CustomerName = "Weir Minerals Netherlands",
                AdministrationNumber = "1",
                DebtorNumber = "1001",
                RelationNumber = "1001",
                IsHighPriority = true,
                SpecialParsingMethod = "WeirAdvanced",
                OrderKeywords = { "purchase order", "po #", "po:", "order number", "order #", "purchase", "buy" },
                QuoteKeywords = { "rfq", "request for quote", "quotation", "quote request", "pricing", "bid" },
                RevisionKeywords = { "revision", "rev ", "update", "change", "modify", "amendment" },
                EmailKeywords = { "weir", "minerals", "netherlands", "nederland" }
            },

            ["weirminerals.coupahost.com"] = new DomainConfig
            {
                Domain = "weirminerals.coupahost.com",
                CustomerName = "Weir Minerals Netherlands",
                AdministrationNumber = "1",
                DebtorNumber = "1001",
                RelationNumber = "1001",
                IsHighPriority = true,
                SpecialParsingMethod = "WeirCoupa",
                OrderKeywords = { "coupa", "purchase order", "po number", "order confirmation" },
                QuoteKeywords = { "sourcing event", "rfp", "rfq", "bid request" },
                RevisionKeywords = { "po change", "order change", "revision" },
                EmailKeywords = { "weir", "coupa", "sourcing", "procurement" }
            },

            // ASW Groep
            ["aswgroep.nl"] = new DomainConfig
            {
                Domain = "aswgroep.nl",
                CustomerName = "ASW Groep",
                AdministrationNumber = "1",
                DebtorNumber = "1002",
                RelationNumber = "1002",
                OrderKeywords = { "bestelling", "inkooporder", "order", "aanvraag" },
                QuoteKeywords = { "offerte", "prijsaanvraag", "quotatie" },
                RevisionKeywords = { "wijziging", "aanpassing", "herziening" },
                EmailKeywords = { "asw", "groep" }
            },

            // Internal Elcotec
            ["elcotec.nl"] = new DomainConfig
            {
                Domain = "elcotec.nl",
                CustomerName = "Elcotec BV Internal",
                AdministrationNumber = "1",
                DebtorNumber = "1000",
                RelationNumber = "1000",
                SpecialParsingMethod = "OutlookForward",
                OrderKeywords = { "internal order", "workshop order", "production" },
                QuoteKeywords = { "internal quote", "cost estimate" },
                RevisionKeywords = { "internal revision", "update" },
                EmailKeywords = { "elcotec", "internal" }
            },

            // GM Simons
            ["gm-simons.be"] = new DomainConfig
            {
                Domain = "gm-simons.be",
                CustomerName = "GM Simons",
                AdministrationNumber = "1",
                DebtorNumber = "1003",
                RelationNumber = "1003",
                OrderKeywords = { "commande", "order", "bestelling", "opdracht" },
                QuoteKeywords = { "devis", "offre", "quotation", "prijsopgave" },
                RevisionKeywords = { "révision", "modification", "wijziging" },
                EmailKeywords = { "simons", "gm" }
            }
        };

        #endregion

        #region Enhanced Email Classification

        /// <summary>
        /// Enhanced email classification with domain-specific intelligence
        /// </summary>
        public static EmailContentType ClassifyEmailWithDomainIntelligence(
            string subject,
            string body,
            string senderEmail,
            out DomainConfig matchedDomain)
        {
            matchedDomain = null;

            try
            {
                var domain = ExtractDomainFromEmail(senderEmail);
                matchedDomain = GetDomainConfig(domain);

                LogDomain($"🔍 Classifying email from domain: {domain}");
                LogDomain($"📧 Subject: {subject?.Substring(0, Math.Min(subject?.Length ?? 0, 50))}...");

                if (matchedDomain != null)
                {
                    LogDomain($"✅ Found domain config for: {matchedDomain.CustomerName}");
                    return ClassifyWithDomainSpecificRules(subject, body, matchedDomain);
                }
                else
                {
                    LogDomain($"⚠️ No specific domain config found, using generic classification");
                    return EmailContentAnalyzer.AnalyzeEmailContent(subject, body, domain);
                }
            }
            catch (Exception ex)
            {
                LogDomain($"❌ Error in domain classification: {ex.Message}");
                return EmailContentType.Unknown;
            }
        }

        /// <summary>
        /// Domain-specific classification using customer-specific rules
        /// </summary>
        /// </summary>
        private static EmailContentType ClassifyWithDomainSpecificRules(
            string subject,
            string body,
            DomainConfig domainConfig)
        {
            var subjectLower = subject?.ToLower() ?? "";
            var bodyLower = body?.ToLower() ?? "";

            LogDomain($"🎯 Using domain-specific rules for: {domainConfig.CustomerName}");

            // High priority check for critical customers
            if (domainConfig.IsHighPriority)
            {
                LogDomain($"⭐ High priority customer detected: {domainConfig.CustomerName}");
            }

            // 1. Revision Detection (highest priority)
            if (ContainsAnyKeyword(subjectLower, domainConfig.RevisionKeywords) ||
                ContainsAnyKeyword(bodyLower, domainConfig.RevisionKeywords))
            {
                LogDomain($"🔄 REVISION detected using domain-specific keywords");
                return EmailContentType.Revision;
            }

            // 2. Order Detection
            if (ContainsAnyKeyword(subjectLower, domainConfig.OrderKeywords) ||
                ContainsAnyKeyword(bodyLower, domainConfig.OrderKeywords))
            {
                LogDomain($"📦 ORDER detected using domain-specific keywords");
                return EmailContentType.Order;
            }

            // 3. Quote Detection
            if (ContainsAnyKeyword(subjectLower, domainConfig.QuoteKeywords) ||
                ContainsAnyKeyword(bodyLower, domainConfig.QuoteKeywords))
            {
                LogDomain($"💰 QUOTE detected using domain-specific keywords");
                return EmailContentType.Quote;
            }

            // 4. Special Weir Processing
            if (domainConfig.Domain.Contains("weir"))
            {
                return ClassifyWeirSpecificContent(subject, body, domainConfig);
            }

            LogDomain($"❓ No specific classification found, returning Unknown");
            return EmailContentType.Unknown;
        }

        /// <summary>
        /// Special classification logic for Weir domains
        /// </summary>
        private static EmailContentType ClassifyWeirSpecificContent(
            string subject,
            string body,
            DomainConfig domainConfig)
        {
            LogDomain($"🔧 Applying Weir-specific classification logic");

            var subjectLower = subject?.ToLower() ?? "";
            var bodyLower = body?.ToLower() ?? "";

            // Weir-specific patterns
            if (Regex.IsMatch(subjectLower, @"po[\s\-#]*\d+", RegexOptions.IgnoreCase))
            {
                LogDomain($"📦 Weir PO pattern detected in subject");
                return EmailContentType.Order;
            }

            if (Regex.IsMatch(subjectLower, @"rfq[\s\-#]*\d+", RegexOptions.IgnoreCase))
            {
                LogDomain($"💰 Weir RFQ pattern detected in subject");
                return EmailContentType.Quote;
            }

            if (subjectLower.Contains("coupa") && subjectLower.Contains("sourcing"))
            {
                LogDomain($"💰 Weir Coupa sourcing event detected");
                return EmailContentType.Quote;
            }

            // Check for specific Weir HTML patterns
            if (body.Contains("order_lines") || body.Contains("quote_lines") || body.Contains("rfq_lines"))
            {
                if (body.Contains("order_lines"))
                {
                    LogDomain($"📦 Weir order HTML table detected");
                    return EmailContentType.Order;
                }
                else
                {
                    LogDomain($"💰 Weir quote HTML table detected");
                    return EmailContentType.Quote;
                }
            }

            return EmailContentType.Unknown;
        }

        #endregion

        #region Enhanced Customer Information Extraction

        /// <summary>
        /// Get customer information with enhanced domain mapping
        /// </summary>
        public static CustomerInfo GetCustomerInfoForDomain(string emailAddress)
        {
            try
            {
                var domain = ExtractDomainFromEmail(emailAddress);
                var domainConfig = GetDomainConfig(domain);

                if (domainConfig != null)
                {
                    LogDomain($"✅ Found customer info for {domain}: {domainConfig.CustomerName}");

                    return new CustomerInfo
                    {
                        CustomerName = domainConfig.CustomerName,
                        AdministrationNumber = domainConfig.AdministrationNumber,
                        DebtorNumber = domainConfig.DebtorNumber,
                        RelationNumber = domainConfig.RelationNumber,
                        EmailDomain = domain,
                        IsHighPriority = domainConfig.IsHighPriority
                    };
                }
                else
                {
                    LogDomain($"⚠️ No customer config found for {domain}, using fallback");
                    return GetFallbackCustomerInfo(domain);
                }
            }
            catch (Exception ex)
            {
                LogDomain($"❌ Error getting customer info: {ex.Message}");
                return GetFallbackCustomerInfo(emailAddress);
            }
        }

        /// <summary>
        /// Determine the best HTML parsing method for the domain
        /// </summary>
        public static string GetOptimalParsingMethod(string emailAddress)
        {
            var domain = ExtractDomainFromEmail(emailAddress);
            var domainConfig = GetDomainConfig(domain);

            if (domainConfig?.SpecialParsingMethod != null)
            {
                LogDomain($"🎯 Using special parsing method: {domainConfig.SpecialParsingMethod} for {domain}");
                return domainConfig.SpecialParsingMethod;
            }

            // Default logic
            if (domain.Contains("weir"))
                return "WeirDomain";
            else if (domain.Contains("elcotec"))
                return "OutlookForward";
            else
                return "Generic";
        }

        #endregion

        #region Enhanced Article Code Extraction

        /// <summary>
        /// Enhanced article code extraction with domain-specific patterns
        /// </summary>
        public static List<string> ExtractArticleCodesWithDomainLogic(
            string subject,
            string body,
            DomainConfig domainConfig)
        {
            var articleCodes = new List<string>();

            try
            {
                LogDomain($"🔍 Extracting article codes for: {domainConfig?.CustomerName ?? "Unknown"}");

                // Standard patterns
                var standardPatterns = new[]
                {
                    @"\b\d{3}\.\d{3}\.\d{4}\b",         // 897.010.1478
                    @"\b\d{3}-\d{3}-\d{4}\b",           // 897-010-1478
                    @"\b\d{6}\.\d{4}\b",                // 123456.7890
                    @"\b[A-Z]{2,3}\d{3,6}\b"            // AB1234, ABC123456
                };

                // Domain-specific patterns
                if (domainConfig?.Domain.Contains("weir") == true)
                {
                    // Weir-specific article code patterns
                    var weirPatterns = new[]
                    {
                        @"\bWEIR[\-\s]*\d{4,8}\b",       // WEIR-123456
                        @"\b[A-Z]{1,3}\d{6,8}\b",        // A1234567
                        @"\b\d{4}\-\d{4}\-\d{4}\b"       // 1234-5678-9012
                    };

                    foreach (var pattern in weirPatterns)
                    {
                        ExtractWithPattern(subject + " " + body, pattern, articleCodes);
                    }
                }

                // Apply standard patterns
                foreach (var pattern in standardPatterns)
                {
                    ExtractWithPattern(subject + " " + body, pattern, articleCodes);
                }

                LogDomain($"✅ Extracted {articleCodes.Count} article codes");
                return articleCodes.Distinct().ToList();
            }
            catch (Exception ex)
            {
                LogDomain($"❌ Error extracting article codes: {ex.Message}");
                return articleCodes;
            }
        }

        #endregion

        #region Utility Methods

        private static void ExtractWithPattern(string text, string pattern, List<string> results)
        {
            var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
            foreach (Match match in matches)
            {
                var code = match.Value.Trim();
                if (!string.IsNullOrEmpty(code) && !results.Contains(code))
                {
                    results.Add(code);
                    LogDomain($"   📝 Found article code: {code}");
                }
            }
        }

        private static bool ContainsAnyKeyword(string text, List<string> keywords)
        {
            return keywords.Any(keyword => text.Contains(keyword.ToLower()));
        }

        private static string ExtractDomainFromEmail(string emailAddress)
        {
            try
            {
                if (string.IsNullOrEmpty(emailAddress)) return "";

                var atIndex = emailAddress.IndexOf('@');
                if (atIndex > 0 && atIndex < emailAddress.Length - 1)
                {
                    return emailAddress.Substring(atIndex + 1).ToLower();
                }
                return emailAddress.ToLower();
            }
            catch
            {
                return "";
            }
        }

        private static DomainConfig GetDomainConfig(string domain)
        {
            if (string.IsNullOrEmpty(domain)) return null;

            // Direct match first
            if (_domainMappings.ContainsKey(domain))
                return _domainMappings[domain];

            // Partial match for subdomains
            foreach (var mapping in _domainMappings)
            {
                if (domain.Contains(mapping.Key) || mapping.Key.Contains(domain))
                {
                    LogDomain($"📍 Partial domain match: {domain} → {mapping.Key}");
                    return mapping.Value;
                }
            }

            return null;
        }

        private static CustomerInfo GetFallbackCustomerInfo(string emailOrDomain)
        {
            return new CustomerInfo
            {
                CustomerName = ConfigurationManager.AppSettings["FallbackCustomer:CustomerName"] ?? "Unknown Customer",
                AdministrationNumber = ConfigurationManager.AppSettings["FallbackCustomer:AdmiNum"] ?? "1",
                DebtorNumber = ConfigurationManager.AppSettings["FallbackCustomer:DebiNum"] ?? "9999",
                RelationNumber = ConfigurationManager.AppSettings["FallbackCustomer:RelaNum"] ?? "9999",
                EmailDomain = ExtractDomainFromEmail(emailOrDomain),
                IsHighPriority = false
            };
        }

        #endregion

        #region Logging

        private static List<string> _domainLog = new List<string>();

        public static List<string> GetDomainLog() => new List<string>(_domainLog);
        public static void ClearDomainLog() => _domainLog.Clear();

        private static void LogDomain(string message)
        {
            var logEntry = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
            _domainLog.Add(logEntry);
            Console.WriteLine($"[DOMAIN_PROCESSOR] {logEntry}");
        }

        #endregion

        #region Domain Configuration Management

        /// <summary>
        /// Load domain mappings from App.config and merge with hardcoded ones
        /// </summary>
        public static void LoadDomainMappingsFromConfig()
        {
            try
            {
                var configMappings = ConfigurationManager.AppSettings["CustomerDomainMappings"];
                if (string.IsNullOrEmpty(configMappings)) return;

                LogDomain($"📋 Loading domain mappings from configuration...");

                var mappingPairs = configMappings.Split(';');
                foreach (var pair in mappingPairs)
                {
                    var parts = pair.Split(':');
                    if (parts.Length != 2) continue;

                    var domain = parts[0].Trim();
                    var customerData = parts[1].Split(',');

                    if (customerData.Length >= 4)
                    {
                        var config = new DomainConfig
                        {
                            Domain = domain,
                            AdministrationNumber = customerData[0].Trim(),
                            DebtorNumber = customerData[1].Trim(),
                            RelationNumber = customerData[2].Trim(),
                            CustomerName = customerData[3].Trim()
                        };

                        // Don't override existing detailed configurations
                        if (!_domainMappings.ContainsKey(domain))
                        {
                            _domainMappings[domain] = config;
                            LogDomain($"   ✅ Added config mapping: {domain} → {config.CustomerName}");
                        }
                    }
                }

                LogDomain($"✅ Loaded {_domainMappings.Count} total domain mappings");
            }
            catch (Exception ex)
            {
                LogDomain($"❌ Error loading domain mappings: {ex.Message}");
            }
        }
        public static bool IsWeirCustomer(CustomerInfo customerInfo)
        {
            if (customerInfo == null) return false;

            return customerInfo.CustomerName.Contains("Weir", StringComparison.OrdinalIgnoreCase) ||
                   customerInfo.EmailDomain.Contains("weir", StringComparison.OrdinalIgnoreCase);
        }
        /// <summary>
        /// Get all configured domains for debugging
        /// </summary>
        public static Dictionary<string, string> GetAllDomainMappings()
        {
            return _domainMappings.ToDictionary(
                kvp => kvp.Key,
                kvp => $"{kvp.Value.CustomerName} ({kvp.Value.DebtorNumber})"
            );
        }

        #endregion
    }

}