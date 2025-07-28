// NEW: ArticleCodeExtractor.cs - Centralized article code extraction and validation
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Mkg_Elcotec_Automation.Utilities
{
    /// <summary>
    /// Centralized article code extraction and validation for Weir domain
    /// Used by OrderLogicService, QuoteLogicService, and RevisionLogicService
    /// </summary>
    public static class ArticleCodeExtractor
    {
        /// <summary>
        /// Extract valid Weir article codes from text content
        /// </summary>
        public static List<string> ExtractArticleCodes(string text)
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
                    @"\b(\d{3}\.\d{2,3}\.\d{2,3})\b",               // 500.143.527
                    @"\b(\d{3}-\d{2}-\d{3})\b",                      // 475-25-041
                    @"\b(\d{3}\.\d{3}\.\d{4}[A-Z]*)\b",             // 897.010.1478
                    @"\b(\d{3}\.\d{2}\.\d{2})\b",                   // 891.75.61
                    @"\b(\d{3}\.\d{2}\.\d{2}\.\d{3})\b",            // 891.75.61.000
                    @"(?:article|part|item)[:\s#]*([A-Z0-9\.\-]{6,})", // Article: ABC123
                    @"(?:sku|pn|part\s*number)[:\s#]*([A-Z0-9\.\-]{6,})" // SKU: ABC123
                };

                foreach (var pattern in patterns)
                {
                    var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
                    foreach (Match match in matches)
                    {
                        var code = match.Groups[1].Value.ToUpper().Trim();

                        // Clean and validate the article code
                        code = CleanArticleCode(code);

                        if (IsValidWeirArticleCode(code))
                        {
                            articleCodes.Add(code);
                            Console.WriteLine($"📦 Found valid article code: {code}");
                        }
                        else
                        {
                            Console.WriteLine($"❌ Rejected invalid code: {code}");
                        }
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

        /// <summary>
        /// Clean article code by removing common artifacts
        /// </summary>
        private static string CleanArticleCode(string code)
        {
            if (string.IsNullOrEmpty(code)) return code;

            // Remove common artifacts appended to article codes
            var artifactsToRemove = new[]
            {
                "DELIVER", "DELIVERY", "QTY", "QUANTITY", "PCS", "PIECES",
                "EA", "EACH", "ST", "STUKS", "UNIT", "UNITS"
            };

            foreach (var artifact in artifactsToRemove)
            {
                if (code.EndsWith(artifact, StringComparison.OrdinalIgnoreCase))
                {
                    code = code.Substring(0, code.Length - artifact.Length);
                    Console.WriteLine($"🧹 Cleaned artifact '{artifact}' from article code");
                }
            }

            return code.Trim();
        }

        /// <summary>
        /// Validate if this is a real Weir article code
        /// </summary>
        private static bool IsValidWeirArticleCode(string code)
        {
            if (string.IsNullOrEmpty(code) || code.Length < 3)
                return false;

            // BLACKLIST: Known invalid codes that are artifacts
            var invalidCodes = new[]
            {
                "NUMBER", "ITEM", "ARTICLE", "CODE", "PART", "DESCRIPTION",
                "QTY", "QUANTITY", "PRICE", "TOTAL", "LINE", "NL002",
                "CAPACITIVESUBCONTRACTING", "MAIL", "WEIR", "VF255"
            };

            if (invalidCodes.Contains(code.ToUpper()))
            {
                Console.WriteLine($"❌ Blacklisted code: {code}");
                return false;
            }

            // WHITELIST: Valid Weir article code patterns (NUMERIC DOT-SEPARATED ONLY)
            var validPatterns = new[]
            {
                @"^\d{3}\.\d{3}\.\d{1,4}$",           // 870.001.272, 897.010.1478
                @"^\d{3}\.\d{3}\.\d{1,4}[A-Z]{1,2}$", // 870.001.272A, 891.75.02.000A1
                @"^\d{3}\.\d{2,3}\.\d{2,4}$",         // 816.310.136
                @"^\d{3}\.\d{3}\.\d{3}\.\d{1,3}$",    // Four-part codes like 891.75.61.000
                @"^\d{3}\.\d{2}\.\d{2}$"              // Three-part codes like 891.75.61
            };

            foreach (var pattern in validPatterns)
            {
                if (Regex.IsMatch(code, pattern))
                {
                    Console.WriteLine($"✅ Valid article code: {code}");
                    return true;
                }
            }

            Console.WriteLine($"❌ No valid pattern matched: {code}");
            return false;
        }

        /// <summary>
        /// Extract article codes from email subject line
        /// </summary>
        public static List<string> ExtractFromSubject(string subject)
        {
            if (string.IsNullOrEmpty(subject))
                return new List<string>();

            // Subject lines often contain article codes in specific formats
            var subjectPatterns = new[]
            {
                @"\b(\d{3}\.\d{2}\.\d{2}(?:\.\d{3})?[A-Z]*)\b", // 891.75.01.000A1
                @"\b(\d{3}\.\d{3}\.\d{1,4})\b"                   // Standard format
            };

            var codes = new HashSet<string>();
            foreach (var pattern in subjectPatterns)
            {
                var matches = Regex.Matches(subject, pattern);
                foreach (Match match in matches)
                {
                    var code = CleanArticleCode(match.Groups[1].Value.ToUpper().Trim());
                    if (IsValidWeirArticleCode(code))
                    {
                        codes.Add(code);
                    }
                }
            }

            return codes.ToList();
        }
    }
}