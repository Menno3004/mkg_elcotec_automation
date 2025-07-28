using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Mkg_Elcotec_Automation.Models;

namespace Mkg_Elcotec_Automation.Services
{
    public static class EmailContentAnalyzer
    {
        private static List<string> DebugLog = new List<string>();

        public static void ClearDebugLog() => DebugLog.Clear();
        public static List<string> GetDebugLog() => new List<string>(DebugLog);

        private static void LogDebug(string message)
        {
            var logEntry = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
            DebugLog.Add(logEntry);
            Console.WriteLine($"[EMAIL_ANALYZER] {logEntry}");
        }

        /// <summary>
        /// Enhanced email classification with better quote vs order detection
        /// FIXED: Now properly detects Weir purchase orders and other business content
        /// </summary>
        public static EmailContentType AnalyzeEmailContent(string subject, string body, string senderDomain)
        {
            LogDebug($"=== ANALYZING EMAIL CONTENT ===");
            LogDebug($"Subject: '{subject?.Substring(0, Math.Min(subject?.Length ?? 0, 100))}'");
            LogDebug($"Body length: {body?.Length ?? 0}");
            LogDebug($"Sender domain: '{senderDomain}'");

            if (string.IsNullOrWhiteSpace(subject) && string.IsNullOrWhiteSpace(body))
            {
                LogDebug("Result: NO BUSINESS CONTENT (empty subject and body)");
                return EmailContentType.NoBusinessContent;
            }

            var subjectLower = subject?.ToLower() ?? "";
            var bodyLower = body?.ToLower() ?? "";

            // PRIORITY 1: REVISION DETECTION (highest priority - very specific)
            if (IsRevisionEmail(subjectLower, bodyLower))
            {
                LogDebug("Result: REVISION EMAIL detected");
                return EmailContentType.Revision;
            }

            // PRIORITY 2: ORDER DETECTION (Enhanced Weir support)
            if (IsOrderEmail(subjectLower, bodyLower, senderDomain))
            {
                LogDebug("Result: ORDER EMAIL detected");
                return EmailContentType.Order;
            }

            // PRIORITY 3: QUOTE DETECTION (RFQ = Request for Quote)
            if (IsQuoteEmail(subjectLower, bodyLower, senderDomain))
            {
                LogDebug("Result: QUOTE EMAIL detected");
                return EmailContentType.Quote;
            }

            LogDebug("Result: NO BUSINESS CONTENT (no patterns matched)");
            return EmailContentType.NoBusinessContent;
        }

        /// <summary>
        /// Enhanced revision detection - very specific criteria
        /// </summary>
        private static bool IsRevisionEmail(string subject, string body)
        {
            LogDebug("--- REVISION DETECTION ---");

            var subjectLower = subject?.ToLower() ?? "";
            var bodyLower = body?.ToLower() ?? "";

            // MINIMAL: Only the most specific revision keywords
            var revisionKeywords = new[]
            {
        "drawing revision",      // Very specific - technical drawing change
        "tekening revisie",      // Dutch: drawing revision
        "drawing changed to rev", // Very specific pattern
        "nieuwe revisie",        // Dutch: new revision (specific)
        "revision change"        // Explicit revision change
    };

            // Check for minimal revision keywords
            foreach (var keyword in revisionKeywords)
            {
                if (subjectLower.Contains(keyword) || bodyLower.Contains(keyword))
                {
                    LogDebug($"✅ Revision keyword found: {keyword}");
                    return true;
                }
            }

            // MINIMAL: Only check for explicit revision number patterns
            var revisionPatterns = new[]
            {
        @"rev\.?\s*[A-Z]\s*(?:to|→|->|naar)\s*rev\.?\s*[A-Z]",  // Rev A to Rev B
        @"changed\s*to\s*rev\.?\s*[A-Z0-9]+"                    // Changed to Rev B
    };

            foreach (var pattern in revisionPatterns)
            {
                if (Regex.IsMatch(subjectLower + " " + bodyLower, pattern, RegexOptions.IgnoreCase))
                {
                    LogDebug($"✅ Revision pattern found: {pattern}");
                    return true;
                }
            }

            LogDebug("❌ No revision indicators found");
            return false;
        }

        /// <summary>
        /// ENHANCED ORDER DETECTION - Fixed for Weir emails
        /// </summary>
        private static bool IsOrderEmail(string subject, string body, string senderDomain)
        {
            LogDebug("--- ORDER DETECTION ---");
            LogDebug($"Checking domain: '{senderDomain}'");

            // WEIR-SPECIFIC DETECTION (highest priority)
            if (senderDomain.Contains("weirminerals.coupahost.com") ||
                senderDomain.Contains("mail.weir") ||
                senderDomain.Contains("weir"))
            {
                LogDebug("🔧 WEIR DOMAIN DETECTED - Using Weir-specific logic");

                // Weir purchase order patterns
                if (subject.Contains("purchase order"))
                {
                    LogDebug("✅ WEIR: 'purchase order' found in subject");
                    return true;
                }

                if (subject.Contains("weir") && subject.Contains("order"))
                {
                    LogDebug("✅ WEIR: 'weir' + 'order' pattern found");
                    return true;
                }

                // Weir order number patterns (#4501508414)
                if (Regex.IsMatch(subject, @"#\d{10}", RegexOptions.IgnoreCase))
                {
                    LogDebug("✅ WEIR: 10-digit order number pattern found");
                    return true;
                }

                // Weir coupa patterns
                if (subject.Contains("coupa") && (subject.Contains("order") || subject.Contains("po")))
                {
                    LogDebug("✅ WEIR: Coupa order pattern found");
                    return true;
                }

                // Check body for Weir order HTML patterns
                if (body.Contains("order_lines") || body.Contains("po_info"))
                {
                    LogDebug("✅ WEIR: Order HTML structure found in body");
                    return true;
                }
            }

            // GENERAL ORDER KEYWORDS (for all domains)
            var orderKeywords = new[]
            {
                "purchase order",
                "po #",
                "po:",
                "order #",
                "order number",
                "inkooporder",       // Dutch purchase order
                "bestelling",        // Dutch order
                "bestelbevestiging", // Dutch order confirmation
                "order confirmation",
                "nieuwe inkooporder", // Dutch: new purchase order
                "inkooporder uitgegeven" // Dutch: purchase order issued
            };

            // Check subject for order keywords
            foreach (var keyword in orderKeywords)
            {
                if (subject.Contains(keyword))
                {
                    LogDebug($"✅ Order keyword found in subject: '{keyword}'");
                    return true;
                }
            }

            // Order number patterns (general)
            var orderNumberPatterns = new[]
            {
                @"po[\s\-#]*\d+",           // PO 123456, PO#123456, PO-123456
                @"order[\s\-#]*\d+",        // Order 123456
                @"bestelling[\s\-#]*\d+",   // Dutch: Bestelling 123456
                @"#\d{4,}",                 // #123456 (4+ digits)
                @"\b\d{8,10}\b"             // 8-10 digit numbers (common for order numbers)
            };

            foreach (var pattern in orderNumberPatterns)
            {
                if (Regex.IsMatch(subject, pattern, RegexOptions.IgnoreCase))
                {
                    LogDebug($"✅ Order number pattern found: {pattern}");
                    return true;
                }
            }

            // Check body for order HTML structures
            if (body.Contains("order_lines") && !body.Contains("quote_lines"))
            {
                LogDebug("✅ Order HTML table structure found in body");
                return true;
            }

            LogDebug("❌ No order indicators found");
            return false;
        }

        /// <summary>
        /// Enhanced quote detection
        /// </summary>
        private static bool IsQuoteEmail(string subject, string body, string senderDomain)
        {
            LogDebug("--- QUOTE DETECTION ---");

            // Strong quote indicators in subject
            var quoteKeywords = new[]
            {
                "rfq",               // Request for Quote
                "request for quote",
                "quote request",
                "quotation",
                "offerte",           // Dutch for quote
                "offerteaanvraag",   // Dutch quote request
                "aanvraag offerte",  // Dutch request for quote
                "pricing request",
                "price request",
                "prijsopgave",       // Dutch price quote
                "prijsaanvraag",     // Dutch price request
                "sourcing event",    // Weir Coupa sourcing
                "bid request"
            };

            // Check subject for quote keywords
            foreach (var keyword in quoteKeywords)
            {
                if (subject.Contains(keyword))
                {
                    LogDebug($"✅ Quote keyword found in subject: '{keyword}'");

                    // Make sure it's NOT actually a purchase order
                    if (!subject.Contains("purchase order") &&
                        !subject.Contains("po #") &&
                        !subject.Contains("inkooporder"))
                    {
                        LogDebug("✅ Confirmed as quote (not purchase order)");
                        return true;
                    }
                    else
                    {
                        LogDebug("❌ Contains purchase order keywords - likely an order");
                    }
                }
            }

            // Check for quote-specific HTML content
            if (body.Contains("quote_lines") || body.Contains("rfq_lines"))
            {
                LogDebug("✅ Quote HTML table structure found");
                return true;
            }

            // Check for RFQ numbers in subject
            var rfqPattern = @"rfq[:\-\s#]*(\d+)";
            if (Regex.IsMatch(subject, rfqPattern, RegexOptions.IgnoreCase))
            {
                LogDebug("✅ RFQ number pattern found in subject");
                return true;
            }

            // Weir-specific quote detection
            if (senderDomain.Contains("weir"))
            {
                if (subject.Contains("request") && !subject.Contains("purchase order"))
                {
                    LogDebug("✅ Weir quote request detected");
                    return true;
                }

                if (subject.Contains("coupa") && subject.Contains("sourcing"))
                {
                    LogDebug("✅ Weir Coupa sourcing event detected");
                    return true;
                }
            }

            LogDebug("❌ No quote indicators found");
            return false;
        }

        /// <summary>
        /// Extract domain from email address
        /// </summary>
        public static string ExtractDomainFromEmail(string emailAddress)
        {
            if (string.IsNullOrEmpty(emailAddress)) return "";

            try
            {
                int atIndex = emailAddress.IndexOf('@');
                if (atIndex == -1) return "";

                return emailAddress.Substring(atIndex + 1).ToLower().Trim();
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Get classification summary for debugging
        /// </summary>
        public static string GetClassificationSummary(string subject, string body, string senderEmail)
        {
            var domain = ExtractDomainFromEmail(senderEmail);
            var contentType = AnalyzeEmailContent(subject, body, domain);

            return $"Email: {subject?.Substring(0, Math.Min(subject?.Length ?? 0, 50))}... " +
                   $"| Domain: {domain} | Classification: {contentType}";
        }
    }

    /// <summary>
    /// Email content type classification
    /// </summary>
    public enum EmailContentType
    {
        NoBusinessContent,
        Order,
        Quote,
        Revision,
        Mixed,
        Unknown
    }
}