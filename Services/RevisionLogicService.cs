// ===============================================
// UPDATED RevisionLogicService.cs - Using ArticleCodeExtractor
// MINIMAL CHANGES: Only replaced article code extraction with ArticleCodeExtractor
// ===============================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Mkg_Elcotec_Automation.Models;
using System.Threading.Tasks;
using Microsoft.Graph.Models;
using Mkg_Elcotec_Automation.Utilities; // 🆕 ADD THIS USING

namespace Mkg_Elcotec_Automation.Services
{
    public static class RevisionLogicService
    {
        private static List<string> processingLog = new List<string>();

        #region Public Properties
        public static List<string> GetProcessingLog() => new List<string>(processingLog);
        public static void ClearProcessingLog() => processingLog.Clear();
        #endregion

        #region Core Revision Business Logic

        /// <summary>
        /// Process and validate revisions with business rules applied
        /// Input: Raw revisions from EmailContentAnalyzer
        /// Output: Business-validated and processed revisions ready for MKG injection
        /// </summary>
        public static List<RevisionLine> ProcessRevisionsForMkgInjection(List<RevisionLine> rawRevisions, string customerDomain = null)
        {
            try
            {
                processingLog.Add($"=== PROCESSING {rawRevisions.Count} REVISIONS FOR MKG INJECTION ===");

                var processedRevisions = new List<RevisionLine>();

                foreach (var revision in rawRevisions)
                {
                    try
                    {
                        // Apply business validation
                        if (!ValidateRevisionBusinessRules(revision))
                        {
                            processingLog.Add($"Revision validation failed for ArtiCode: {revision.ArtiCode}");
                            continue;
                        }

                        // Apply revision business rules
                        ApplyRevisionBusinessRules(revision, customerDomain);

                        // Calculate revision impact
                        CalculateRevisionImpact(revision);

                        // Assign priority and approval requirements
                        AssignRevisionPriorityAndApproval(revision, customerDomain);

                        // Format for MKG standards
                        FormatRevisionForMkgStandards(revision);

                        // Handle file versioning business rules
                        ApplyFileVersioningRules(revision);

                        processedRevisions.Add(revision);
                        processingLog.Add($"✓ Successfully processed revision: {revision.ArtiCode} ({revision.CurrentRevision} → {revision.NewRevision})");
                    }
                    catch (Exception ex)
                    {
                        processingLog.Add($"Error processing revision {revision.ArtiCode}: {ex.Message}");
                    }
                }

                processingLog.Add($"Processed {processedRevisions.Count}/{rawRevisions.Count} revisions successfully");
                return processedRevisions;
            }
            catch (Exception ex)
            {
                processingLog.Add($"Critical error in revision processing: {ex.Message}");
                return new List<RevisionLine>();
            }
        }

        /// <summary>
        /// 🆕 ADDED: Extract revisions from email content using centralized article code extraction
        /// </summary>
        public static List<RevisionLine> ExtractRevisions(string subject, string emailBody, string emailDomain)
        {
            try
            {
                processingLog.Add($"🎯 ExtractRevisions called with domain: {emailDomain}");
                var revisions = new List<RevisionLine>();

                // 🆕 USE CENTRALIZED EXTRACTOR
                var articleCodes = ArticleCodeExtractor.ExtractArticleCodes(subject + " " + emailBody);
                processingLog.Add($"📦 Found {articleCodes.Count} valid article codes: {string.Join(", ", articleCodes)}");

                if (!articleCodes.Any())
                {
                    processingLog.Add("📦 No article codes found for revision extraction");
                    return revisions;
                }

                foreach (var articleCode in articleCodes)
                {
                    // Extract revision information
                    var revisionInfo = ExtractRevisionDetailsInline(subject, emailBody);

                    var revision = new RevisionLine
                    {
                        ArtiCode = articleCode,
                        Description = revisionInfo.Description ?? $"Revision for {articleCode}",
                        CurrentRevision = revisionInfo.CurrentRevision ?? "A0",
                        NewRevision = revisionInfo.NewRevision ?? "A1",
                        RevisionReason = revisionInfo.RevisionReason ?? "Revision from email",
                        Quantity = revisionInfo.Quantity ?? "1",
                        Unit = revisionInfo.Unit ?? "PCS",
                        RevisionDate = DateTime.Now.ToString("yyyy-MM-dd"),
                        RevisionStatus = "Draft",
                        Priority = "Normal",
                        ApprovalRequired = "Yes",
                        ClientDomain = emailDomain,
                        ExtractionMethod = "EMAIL_EXTRACTION",
                        ExtractionTimestamp = DateTime.Now
                    };

                    revisions.Add(revision);
                    processingLog.Add($"✅ Created revision for {articleCode}");
                }

                processingLog.Add($"📦 Total revisions created: {revisions.Count}");
                return revisions;
            }
            catch (Exception ex)
            {
                processingLog.Add($"❌ Error in ExtractRevisions: {ex.Message}");
                return new List<RevisionLine>();
            }
        }

        /// <summary>
        /// 🔄 UPDATED: Extract article codes from email subject using centralized extractor
        /// </summary>
        public static List<string> ExtractArticleCodesFromSubject(string subject)
        {
            // 🆕 USE CENTRALIZED EXTRACTOR instead of local implementation
            return ArticleCodeExtractor.ExtractFromSubject(subject);
        }

        #endregion

        #region Private Helper Methods

        private static (string Description, string CurrentRevision, string NewRevision, string RevisionReason, string Quantity, string Unit) ExtractRevisionDetailsInline(string subject, string emailBody)
        {
            var description = ExtractDescriptionInline(subject, emailBody);
            var (currentRev, newRev) = ExtractRevisionNumbers(subject, emailBody);
            var revisionReason = ExtractRevisionReason(subject, emailBody);
            var quantity = ExtractQuantityInline(subject, emailBody);
            var unit = ExtractUnitInline(emailBody);

            return (description, currentRev, newRev, revisionReason, quantity, unit);
        }

        private static string ExtractDescriptionInline(string subject, string emailBody)
        {
            if (!string.IsNullOrEmpty(subject) && subject.Length > 10)
                return subject.Substring(0, Math.Min(subject.Length, 100));

            return "Revision extracted from email";
        }

        private static (string currentRevision, string newRevision) ExtractRevisionNumbers(string subject, string emailBody)
        {
            var text = subject + " " + emailBody;

            // Pattern for revision changes like "A0 to A1" or "Rev A → Rev B"
            var revisionPatterns = new[]
            {
                @"(?:rev|revision)[:\s]*([A-Z]\d+)[:\s]*(?:to|→|->)[:\s]*([A-Z]\d+)",
                @"([A-Z]\d+)[:\s]*(?:to|→|->)[:\s]*([A-Z]\d+)",
                @"(?:from|old)[:\s]*([A-Z]\d+)[:\s]*(?:to|new)[:\s]*([A-Z]\d+)"
            };

            foreach (var pattern in revisionPatterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    return (match.Groups[1].Value.ToUpper(), match.Groups[2].Value.ToUpper());
                }
            }

            // Default assumption if no specific revision change found
            return ("A0", "A1");
        }

        private static string ExtractRevisionReason(string subject, string emailBody)
        {
            var reasonPatterns = new[]
            {
                @"reason[:\s]*([^.\n]{10,100})",
                @"change[:\s]*([^.\n]{10,100})",
                @"update[:\s]*([^.\n]{10,100})"
            };

            foreach (var pattern in reasonPatterns)
            {
                var match = Regex.Match(subject + " " + emailBody, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    return match.Groups[1].Value.Trim();
                }
            }

            return "Revision change as per email request";
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

        #endregion

        #region Business Rule Validation

        /// <summary>
        /// Validate revision meets business requirements
        /// </summary>
        public static bool ValidateRevisionBusinessRules(RevisionLine revision, EnhancedProgressManager progressManager = null)
        {
            var errors = new List<string>();

            // Business Rule 1: Article Code is mandatory
            if (string.IsNullOrWhiteSpace(revision.ArtiCode))
                errors.Add("Article Code is required");

            // Business Rule 2: Current revision must be specified
            if (string.IsNullOrWhiteSpace(revision.CurrentRevision))
                errors.Add("Current revision is required");

            // Business Rule 3: New revision must be specified and different
            if (string.IsNullOrWhiteSpace(revision.NewRevision))
                errors.Add("New revision is required");
            else if (revision.CurrentRevision == revision.NewRevision)
                errors.Add("New revision must be different from current revision");

            // Business Rule 4: Revision sequence validation
            if (!string.IsNullOrEmpty(revision.CurrentRevision) && !string.IsNullOrEmpty(revision.NewRevision))
            {
                if (!IsValidRevisionSequence(revision.CurrentRevision, revision.NewRevision))
                    errors.Add("Invalid revision sequence - new revision must follow logical progression");
            }

            // Business Rule 5: Revision reason is mandatory for major changes
            if (string.IsNullOrWhiteSpace(revision.RevisionReason))
                errors.Add("Revision reason is required");

            // Business Rule 6: Valid quantity format (if specified)
            if (!string.IsNullOrWhiteSpace(revision.Quantity))
            {
                if (!decimal.TryParse(revision.Quantity.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal qty) || qty <= 0)
                    errors.Add("Quantity must be a positive number");
            }

            // Business Rule 7: Valid target price (if specified)
            if (!string.IsNullOrEmpty(revision.QuotedPrice))
            {
                if (!decimal.TryParse(revision.QuotedPrice.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                    errors.Add("Target price must be a valid number");
            }

            if (errors.Any())
            {
                processingLog.Add($"Validation errors for {revision.ArtiCode}: {string.Join(", ", errors)}");

                // 🎯 INCREMENT B.ERRORS for pre-validation failures
                progressManager?.IncrementBusinessErrors("Pre-Validation");

                return false;
            }

            return true;
        }

        /// <summary>
        /// Apply customer-specific revision business rules
        /// </summary>
        public static void ApplyRevisionBusinessRules(RevisionLine revision, string customerDomain)
        {
            // Business Rule: Set default values if missing
            if (string.IsNullOrEmpty(revision.Unit))
                revision.Unit = "PCS";

            if (string.IsNullOrEmpty(revision.RevisionStatus))
                revision.RevisionStatus = "Draft";

            if (string.IsNullOrEmpty(revision.Priority))
                revision.Priority = "Normal";

            if (string.IsNullOrEmpty(revision.ApprovalRequired))
                revision.ApprovalRequired = "Yes"; // Default to requiring approval

            // Business Rule: Set revision date
            if (string.IsNullOrEmpty(revision.RevisionDate))
            {
                revision.RevisionDate = DateTime.Now.ToString("yyyy-MM-dd");
            }

            // Business Rule: Auto-increment revision if not specified
            if (string.IsNullOrEmpty(revision.NewRevision) && !string.IsNullOrEmpty(revision.CurrentRevision))
            {
                revision.NewRevision = IncrementRevision(revision.CurrentRevision);
            }

            // Business Rule: Customer-specific processing
            ApplyCustomerSpecificRevisionRules(revision, customerDomain);

            // Business Rule: Set technical/commercial change defaults
            if (string.IsNullOrEmpty(revision.TechnicalChanges))
                revision.TechnicalChanges = "Technical changes as per revision request";

            if (string.IsNullOrEmpty(revision.CommercialChanges))
                revision.CommercialChanges = "No commercial impact";
        }

        /// <summary>
        /// Calculate revision impact on cost and schedule
        /// </summary>
        public static void CalculateRevisionImpact(RevisionLine revision)
        {
            // Business logic for calculating revision impact
            try
            {
                // Log impact level based on revision type for future use
                if (revision.CurrentRevision == "A0" || revision.NewRevision.Contains("A"))
                {
                    processingLog.Add($"Minor revision detected for {revision.ArtiCode}: {revision.CurrentRevision} → {revision.NewRevision}");
                }
                else
                {
                    processingLog.Add($"Major revision detected for {revision.ArtiCode}: {revision.CurrentRevision} → {revision.NewRevision}");
                }
            }
            catch (Exception ex)
            {
                processingLog.Add($"Error calculating revision impact for {revision.ArtiCode}: {ex.Message}");
            }
        }

        /// <summary>
        /// Assign revision priority and approval requirements
        /// </summary>
        public static void AssignRevisionPriorityAndApproval(RevisionLine revision, string customerDomain)
        {
            // Business rules for priority and approval
            if (customerDomain?.Contains("weir") == true)
            {
                revision.Priority = "High"; // Weir gets high priority
                revision.ApprovalRequired = "Yes"; // Always require approval for Weir
            }
            else
            {
                revision.Priority = "Normal";
                revision.ApprovalRequired = "Yes"; // Default to requiring approval
            }
        }

        /// <summary>
        /// Format revision for MKG standards
        /// </summary>
        public static void FormatRevisionForMkgStandards(RevisionLine revision)
        {
            // Ensure proper formatting for MKG system
            revision.CurrentRevision = revision.CurrentRevision?.ToUpper();
            revision.NewRevision = revision.NewRevision?.ToUpper();
        }

        /// <summary>
        /// Apply file versioning rules
        /// </summary>
        public static void ApplyFileVersioningRules(RevisionLine revision)
        {
            // Business rules for file versioning - log for future implementation
            if (!string.IsNullOrEmpty(revision.DrawingNumber))
            {
                var newDrawingNumber = $"{revision.DrawingNumber}_Rev{revision.NewRevision}";
                processingLog.Add($"File versioning for {revision.ArtiCode}: {revision.DrawingNumber} → {newDrawingNumber}");
            }
        }

        #endregion

        #region Utility Methods

        private static bool IsValidRevisionSequence(string currentRevision, string newRevision)
        {
            // Simple validation - new revision should be "greater" than current
            if (string.IsNullOrEmpty(currentRevision) || string.IsNullOrEmpty(newRevision))
                return false;

            return string.Compare(newRevision, currentRevision, StringComparison.OrdinalIgnoreCase) > 0;
        }

        private static string IncrementRevision(string currentRevision)
        {
            if (string.IsNullOrEmpty(currentRevision))
                return "A1";

            // Simple increment logic for revisions like A0 → A1, B0 → B1
            if (currentRevision.Length >= 2)
            {
                var letter = currentRevision[0];
                if (int.TryParse(currentRevision.Substring(1), out int number))
                {
                    return $"{letter}{number + 1}";
                }
            }

            return "A1"; // Default fallback
        }

        private static void ApplyCustomerSpecificRevisionRules(RevisionLine revision, string customerDomain)
        {
            if (customerDomain?.Contains("weir") == true)
            {
                // Weir-specific rules
                revision.Priority = "High";
                revision.ApprovalRequired = "Yes";
            }
        }
        public static async Task<List<RevisionLine>> ExtractRevisionsSafe(string emailBody, string emailDomain, string subject, object attachments)
        {
            try
            {
                await Task.Delay(1);
                return ExtractRevisions(subject, emailBody, emailDomain);
            }
            catch (Exception ex)
            {
                processingLog.Add($"Error in ExtractRevisionsSafe: {ex.Message}");
                return new List<RevisionLine>();
            }
        }
        #endregion
    }
}