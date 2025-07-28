using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Mkg_Elcotec_Automation.Models;
using System.Threading.Tasks;
using Microsoft.Graph.Models;

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
        #region Missing Method Implementations

        /// <summary>
        /// Check if email contains revision content
        /// </summary>
        /// <summary>
        /// Check if email contains revision content - ENHANCED to reduce false positives
        /// </summary>
        public static bool IsRevisionEmail(string subject, string emailBody, AttachmentCollectionResponse attachments)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(subject))
                    return false;

                var subjectLower = subject.ToLower();
                var bodyLower = emailBody?.ToLower() ?? "";

                // FIRST: Exclude obvious order/quote emails (NEW LOGIC)
                var excludeKeywords = new[]
                {
            "purchase order",      // "Weir Minerals Netherlands Purchase Order"
            "inkooporder",        // "Tenfold | Inkooporder IOR2501086"
            "po #", "order #",    // Order number references
            "quote request",      // Quote requests
            "offerte",           // Dutch quote
            "rfq",               // Request for quote
            "order confirmation" // Order confirmations
        };

                // If subject contains obvious order/quote keywords, NOT a revision
                foreach (var excludeWord in excludeKeywords)
                {
                    if (subjectLower.Contains(excludeWord))
                    {
                        Console.WriteLine($"❌ IsRevisionEmail = FALSE for: {subject} (excluded: {excludeWord})");
                        return false;
                    }
                }

                // SECOND: Check for SPECIFIC revision keywords (keeping your original logic but enhanced)
                var revisionKeywords = new[]
                {
            "drawing revision",   // Very specific - technical drawing change
            "tekening revisie",   // Dutch: drawing revision  
            "revised drawing",    // "Revised drawing 897.010.1061G"
            "revision change",    // Explicit revision change
            "nieuwe revisie",     // Dutch: new revision (specific)
            "changed to rev",     // "Changed to rev B"
            "dwg revision"        // Drawing revision short
        };

                // Check for specific revision keywords in subject (most important)
                foreach (var keyword in revisionKeywords)
                {
                    if (subjectLower.Contains(keyword))
                    {
                        Console.WriteLine($"✅ IsRevisionEmail = TRUE for: {subject} (found specific keyword: {keyword})");
                        return true;
                    }
                }

                // THIRD: Check for explicit revision patterns like "rev A to B"
                var revisionPatterns = new[]
                {
            @"rev\.?\s*[A-Z0-9]\s*(?:to|→|->|naar)\s*[A-Z0-9]",  // "rev A to B" or "rev 1 to 2"
            @"revision\s*[A-Z0-9]\s*(?:to|→|->|naar)\s*[A-Z0-9]", // "revision A to B"
            @"revisie\s*[A-Z0-9]\s*(?:to|→|->|naar)\s*[A-Z0-9]",  // Dutch "revisie A naar B"
        };

                var fullContent = subjectLower + " " + bodyLower;
                foreach (var pattern in revisionPatterns)
                {
                    if (Regex.IsMatch(fullContent, pattern, RegexOptions.IgnoreCase))
                    {
                        Console.WriteLine($"✅ IsRevisionEmail = TRUE for: {subject} (found revision pattern)");
                        return true;
                    }
                }

                // FOURTH: Only allow generic keywords if they appear in very specific contexts
                var genericRevisionWords = new[] { "revised", "revision", "revisie", "update", "wijziging" };
                foreach (var word in genericRevisionWords)
                {
                    if (subjectLower.Contains(word))
                    {
                        // Only accept if it's in a technical context (has article codes or drawing references)
                        if (HasTechnicalContext(subject, emailBody))
                        {
                            Console.WriteLine($"✅ IsRevisionEmail = TRUE for: {subject} (generic word '{word}' in technical context)");
                            return true;
                        }
                    }
                }

                Console.WriteLine($"❌ IsRevisionEmail = FALSE for: {subject}");
                return false;
            }
            catch (Exception ex)
            {
                processingLog.Add($"Error in IsRevisionEmail: {ex.Message}");
                Console.WriteLine($"❌ IsRevisionEmail ERROR for: {subject} - {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Helper method to check if email has technical context (NEW METHOD)
        /// </summary>
        private static bool HasTechnicalContext(string subject, string emailBody)
        {
            var content = (subject + " " + emailBody).ToLower();

            // Look for drawing/technical indicators
            var technicalIndicators = new[]
                                                {
                                            "drawing", "dwg", "tekening",
                                            "specification", "material",
                                            "dimension", "tolerance"
                                        };

            // Look for article code patterns
            var hasArticleCodes = Regex.IsMatch(content, @"\d{3}\.\d{3}\.\d{3}", RegexOptions.IgnoreCase);

            // Look for revision number patterns
            var hasRevNumbers = Regex.IsMatch(content, @"rev\.?\s*[A-Z0-9]", RegexOptions.IgnoreCase);

            var hasTechnicalTerms = technicalIndicators.Any(term => content.Contains(term));

            return hasArticleCodes || hasRevNumbers || hasTechnicalTerms;
        }

        /// <summary>
        /// Extract revisions from email content
        /// </summary>

        public static List<RevisionLine> ExtractRevisions(string emailBody, string emailDomain, string subject, AttachmentCollectionResponse attachments)
        {
            try
            {
                processingLog.Add("=== EXTRACTING REVISIONS FROM EMAIL ===");

                var revisions = new List<RevisionLine>();

                // Only process if this is actually a revision email
                if (!IsRevisionEmail(subject, emailBody, attachments))
                {
                    processingLog.Add("❌ Not a revision email - returning empty list");
                    return revisions;
                }

                // Extract article codes from the email
                var articleCodes = ExtractArticleCodesFromSubject(subject);

                if (articleCodes.Count == 0)
                {
                    processingLog.Add("❌ No articles found - returning empty list");
                    return revisions;
                }

                processingLog.Add($"✅ Found {articleCodes.Count} articles for revision");

                // Create simple revisions for each article code
                foreach (var articleCode in articleCodes)
                {
                    var revision = new RevisionLine
                    {
                        ArtiCode = articleCode,
                        DrawingNumber = articleCode,
                        CurrentRevision = "A",
                        NewRevision = "B",
                        RevisionReason = $"Drawing revision from email: {subject}",
                        TechnicalChanges = "Technical revision as per email",
                        CommercialChanges = "",
                        RevisionDate = DateTime.Now.ToString("yyyy-MM-dd"),
                        Priority = "Normal",
                        ApprovalRequired = "Yes",
                        RevisionStatus = "Draft"
                    };

                    revision.SetExtractionDetails("REVISION_EMAIL_EXTRACTION", emailDomain);
                    revisions.Add(revision);

                    processingLog.Add($"✓ Created revision: {articleCode}");
                }

                processingLog.Add($"✅ Extracted {revisions.Count} revisions from revision email");
                return revisions;
            }
            catch (Exception ex)
            {
                processingLog.Add($"❌ Error in ExtractRevisions: {ex.Message}");
                return new List<RevisionLine>();
            }
        }
        /// <summary>
        /// Determine revision priority based on content
        /// </summary>
        public static string DetermineRevisionPriority(string subject, string body)
        {
            var contentLower = (subject + " " + body).ToLower();

            // High priority indicators
            if (contentLower.Contains("urgent") || contentLower.Contains("immediate") ||
                contentLower.Contains("critical") || contentLower.Contains("asap") ||
                contentLower.Contains("spoedig") || contentLower.Contains("dringend"))
            {
                return "High";
            }

            // Low priority indicators
            if (contentLower.Contains("minor") || contentLower.Contains("small") ||
                contentLower.Contains("klein") || contentLower.Contains("minor"))
            {
                return "Low";
            }

            return "Normal";
        }
        /// <summary>
        /// Extract detailed revision information from email content
        /// </summary>
        public static (string DrawingNumber, string CurrentRevision, string NewRevision, string Reason, string TechnicalChanges, string CommercialChanges, string ApprovalRequired) ExtractRevisionDetails(string subject, string body)
        {
            // Extract drawing number
            string drawingNumber = null;
            var dwgPattern = @"(dwg|drawing|tekening)[\s\-:]*([a-zA-Z0-9\-_]+)";
            var dwgMatch = Regex.Match(subject + " " + body, dwgPattern, RegexOptions.IgnoreCase);
            if (dwgMatch.Success)
            {
                drawingNumber = dwgMatch.Groups[2].Value.Trim();
            }

            // Extract revision numbers (from -> to pattern)
            var revPattern = @"rev\.?\s*([A-Z0-9]+).*?(?:to|→|->|naar).*?([A-Z0-9]+)";
            var revMatch = Regex.Match(subject + " " + body, revPattern, RegexOptions.IgnoreCase);

            string currentRev = "A";
            string newRev = "B";

            if (revMatch.Success)
            {
                currentRev = revMatch.Groups[1].Value.ToUpper();
                newRev = revMatch.Groups[2].Value.ToUpper();
            }
            else
            {
                // Try alternative pattern: rev A, rev B
                var altRevPattern = @"rev\.?\s*([A-Z0-9]+)";
                var altMatches = Regex.Matches(subject + " " + body, altRevPattern, RegexOptions.IgnoreCase);
                if (altMatches.Count >= 2)
                {
                    currentRev = altMatches[0].Groups[1].Value.ToUpper();
                    newRev = altMatches[1].Groups[1].Value.ToUpper();
                }
            }

            // Extract reason
            string reason = "Revision update";
            if (body.Contains("reason:") || body.Contains("reden:"))
            {
                var reasonPattern = @"(reason|reden):\s*([^\n\r]+)";
                var reasonMatch = Regex.Match(body, reasonPattern, RegexOptions.IgnoreCase);
                if (reasonMatch.Success)
                {
                    reason = reasonMatch.Groups[2].Value.Trim();
                }
            }

            // Extract technical changes
            string technicalChanges = null;
            if (body.Contains("technical") || body.Contains("technisch"))
            {
                var techPattern = @"(technical|technisch)[^:]*:\s*([^\n\r]+)";
                var techMatch = Regex.Match(body, techPattern, RegexOptions.IgnoreCase);
                if (techMatch.Success)
                {
                    technicalChanges = techMatch.Groups[2].Value.Trim();
                }
            }

            // Extract commercial changes
            string commercialChanges = null;
            if (body.Contains("commercial") || body.Contains("commercieel"))
            {
                var commPattern = @"(commercial|commercieel)[^:]*:\s*([^\n\r]+)";
                var commMatch = Regex.Match(body, commPattern, RegexOptions.IgnoreCase);
                if (commMatch.Success)
                {
                    commercialChanges = commMatch.Groups[2].Value.Trim();
                }
            }

            // Determine if approval is required
            string approvalRequired = "Yes"; // Default to requiring approval
            if (body.ToLower().Contains("no approval") || body.ToLower().Contains("geen goedkeuring"))
            {
                approvalRequired = "No";
            }

            return (drawingNumber, currentRev, newRev, reason, technicalChanges, commercialChanges, approvalRequired);
        }
        /// <summary>
        /// Extract article codes from email subject
        /// </summary>
        public static List<string> ExtractArticleCodesFromSubject(string subject)
        {
            var articleCodes = new List<string>();

            if (string.IsNullOrEmpty(subject))
                return articleCodes;

            // Enhanced patterns for article code detection
            var patterns = new[]{
                                    @"\b(\d{3}\.\d{3}\.\d{3})\b",           // Standard format: 123.456.789
                                    @"\b(\d{3}-\d{3}-\d{3})\b",             // Dash format: 123-456-789
                                    @"\b([A-Z]{2,3}\d{3,6})\b",             // Alpha-numeric: ABC123456
                                    @"\bP/N[\s:]*([A-Z0-9\-\.]{6,})\b",     // Part number: P/N ABC-123-456
                                    @"\bPN[\s:]*([A-Z0-9\-\.]{6,})\b",      // Part number: PN ABC-123-456
                                    @"\bART[\s:]*([A-Z0-9\-\.]{6,})\b"      // Article: ART 123-456-789
                                };

            foreach (var pattern in patterns)
            {
                var matches = Regex.Matches(subject, pattern, RegexOptions.IgnoreCase);
                foreach (Match match in matches)
                {
                    var code = match.Groups[1].Value.Trim();
                    if (!articleCodes.Contains(code))
                    {
                        articleCodes.Add(code);
                    }
                }
            }

            return articleCodes;
        }
        #endregion
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
        /// Calculate revision impact and complexity
        /// </summary>
        public static void CalculateRevisionImpact(RevisionLine revision)
        {
            try
            {
                // Business Rule: Determine revision complexity
                var revisionComplexity = DetermineRevisionComplexity(revision);
                processingLog.Add($"Revision complexity for {revision.ArtiCode}: {revisionComplexity}");

                // Business Rule: Major revisions require additional approval
                if (revisionComplexity == "Major")
                {
                    revision.ApprovalRequired = "Yes";
                    revision.Priority = "High";
                }

                // Business Rule: Calculate pricing impact
                if (!string.IsNullOrEmpty(revision.QuotedPrice))
                {
                    var priceImpact = CalculatePriceImpact(revision);
                    if (priceImpact > 0.10m) // More than 10% price change
                    {
                        revision.ApprovalRequired = "Yes";
                        revision.CommercialChanges = $"Price impact: {priceImpact:P2}";
                    }
                }

                // Business Rule: Timeline impact assessment
                AssessTimelineImpact(revision);
            }
            catch (Exception ex)
            {
                processingLog.Add($"Error calculating revision impact for {revision.ArtiCode}: {ex.Message}");
            }
        }

        /// <summary>
        /// Assign priority and approval requirements based on business rules
        /// </summary>
        public static void AssignRevisionPriorityAndApproval(RevisionLine revision, string customerDomain)
        {
            try
            {
                // Business Rule: High-priority customers get expedited processing
                if (!string.IsNullOrEmpty(customerDomain))
                {
                    var domain = customerDomain.ToLower();
                    if (domain.Contains("weir") || domain.Contains("shell") || domain.Contains("bp"))
                    {
                        revision.Priority = "High";
                        processingLog.Add($"Applied high-priority customer rule to revision {revision.ArtiCode}");
                    }
                }

                // Business Rule: Major revision changes require approval
                if (IsMajorRevisionChange(revision.CurrentRevision, revision.NewRevision))
                {
                    revision.ApprovalRequired = "Yes";
                    revision.Priority = "High";
                    processingLog.Add($"Major revision detected for {revision.ArtiCode} - approval required");
                }

                // Business Rule: Emergency revisions
                if (revision.RevisionReason?.ToLower().Contains("urgent") == true ||
                    revision.RevisionReason?.ToLower().Contains("emergency") == true)
                {
                    revision.Priority = "High";
                    processingLog.Add($"Emergency revision detected for {revision.ArtiCode}");
                }

                // Business Rule: Drawing revisions require technical approval
                if (!string.IsNullOrEmpty(revision.DrawingNumber))
                {
                    revision.ApprovalRequired = "Yes";
                    processingLog.Add($"Drawing revision detected for {revision.ArtiCode} - technical approval required");
                }
            }
            catch (Exception ex)
            {
                processingLog.Add($"Error assigning priority for {revision.ArtiCode}: {ex.Message}");
                revision.Priority = "Normal"; // Fallback
            }
        }

        /// <summary>
        /// Format revision data according to MKG standards
        /// </summary>
        public static void FormatRevisionForMkgStandards(RevisionLine revision)
        {
            // MKG Standard: Article code format
            if (!string.IsNullOrEmpty(revision.ArtiCode))
            {
                revision.ArtiCode = revision.ArtiCode.Trim().ToUpper();
            }

            // MKG Standard: Revision format (always uppercase, padded)
            if (!string.IsNullOrEmpty(revision.CurrentRevision))
            {
                revision.CurrentRevision = FormatRevisionNumber(revision.CurrentRevision);
            }

            if (!string.IsNullOrEmpty(revision.NewRevision))
            {
                revision.NewRevision = FormatRevisionNumber(revision.NewRevision);
            }

            // MKG Standard: Unit standardization
            revision.Unit = StandardizeUnit(revision.Unit);

            // MKG Standard: Date formats (yyyy-MM-dd)
            if (!string.IsNullOrEmpty(revision.RevisionDate))
            {
                if (DateTime.TryParse(revision.RevisionDate, out DateTime date))
                {
                    revision.RevisionDate = date.ToString("yyyy-MM-dd");
                }
            }

            // MKG Standard: Description cleanup
            if (!string.IsNullOrEmpty(revision.Description))
            {
                revision.Description = CleanDescription(revision.Description);
            }

            // MKG Standard: Revision reason formatting
            if (!string.IsNullOrEmpty(revision.RevisionReason))
            {
                revision.RevisionReason = CleanRevisionReason(revision.RevisionReason);
            }
        }

        /// <summary>
        /// Apply file versioning business rules
        /// </summary>
        public static void ApplyFileVersioningRules(RevisionLine revision)
        {
            try
            {
                // Business Rule: Generate file backup requirements
                if (revision.ApprovalRequired == "Yes")
                {
                    processingLog.Add($"File backup required for {revision.ArtiCode} due to approval requirement");
                }

                // Business Rule: Drawing file handling
                if (!string.IsNullOrEmpty(revision.DrawingNumber))
                {
                    var oldDrawingName = $"{revision.DrawingNumber}_{revision.CurrentRevision}";
                    var newDrawingName = $"{revision.DrawingNumber}_{revision.NewRevision}";
                    processingLog.Add($"Drawing file transition: {oldDrawingName} → {newDrawingName}");
                }

                // Business Rule: Archive old revision files
                if (IsMajorRevisionChange(revision.CurrentRevision, revision.NewRevision))
                {
                    processingLog.Add($"Archive requirement for {revision.ArtiCode} - major revision change");
                }
            }
            catch (Exception ex)
            {
                processingLog.Add($"Error applying file versioning rules for {revision.ArtiCode}: {ex.Message}");
            }
        }

        #endregion

        #region Revision Management Functions

        /// <summary>
        /// Filter revisions by business criteria
        /// </summary>
        public static List<RevisionLine> FilterRevisionsByStatus(List<RevisionLine> revisions, string status)
        {
            return revisions.Where(r => r.RevisionStatus.Equals(status, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        public static List<RevisionLine> FilterRevisionsByPriority(List<RevisionLine> revisions, string priority)
        {
            return revisions.Where(r => r.Priority.Equals(priority, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        public static List<RevisionLine> FilterRevisionsByApprovalRequired(List<RevisionLine> revisions)
        {
            return revisions.Where(r => r.ApprovalRequired == "Yes").ToList();
        }

        /// <summary>
        /// Update revision status with business rules
        /// </summary>
        public static void UpdateRevisionStatus(RevisionLine revision, string newStatus)
        {
            var validStatuses = new[] { "Draft", "Pending", "Approved", "Rejected", "Implemented", "Cancelled" };

            if (validStatuses.Contains(newStatus))
            {
                revision.RevisionStatus = newStatus;
                processingLog.Add($"Revision {revision.ArtiCode} status changed to {newStatus}");

                // Business Rule: Status-specific actions
                if (newStatus == "Approved")
                {
                    revision.ApprovalRequired = "No";
                }
            }
            else
            {
                processingLog.Add($"Invalid status '{newStatus}' for revision {revision.ArtiCode}");
            }
        }

        /// <summary>
        /// Generate comprehensive revision summary for business reporting
        /// </summary>
        public static string GenerateRevisionSummary(List<RevisionLine> revisions)
        {
            var summary = new List<string>();

            summary.Add($"=== REVISION BUSINESS SUMMARY ===");
            summary.Add($"Total Revisions: {revisions.Count}");
            summary.Add($"High Priority: {revisions.Count(r => r.Priority == "High")}");
            summary.Add($"Pending Approval: {revisions.Count(r => r.ApprovalRequired == "Yes")}");
            summary.Add($"Major Revisions: {revisions.Count(r => IsMajorRevisionChange(r.CurrentRevision, r.NewRevision))}");
            summary.Add($"Drawing Changes: {revisions.Count(r => !string.IsNullOrEmpty(r.DrawingNumber))}");

            // Group by revision types
            var revisionTypes = revisions.GroupBy(r => DetermineRevisionComplexity(r))
                                       .ToDictionary(g => g.Key, g => g.Count());

            summary.Add($"Revision Types:");
            foreach (var type in revisionTypes)
            {
                summary.Add($"  {type.Key}: {type.Value}");
            }

            return string.Join(Environment.NewLine, summary);
        }

        #endregion

        #region Helper Methods

        private static void ApplyCustomerSpecificRevisionRules(RevisionLine revision, string customerDomain)
        {
            if (string.IsNullOrEmpty(customerDomain)) return;

            var domain = customerDomain.ToLower();

            // Weir-specific rules
            if (domain.Contains("weir"))
            {
                revision.Priority = "High";
                revision.ApprovalRequired = "Yes";
                processingLog.Add($"Applied Weir-specific revision rules to {revision.ArtiCode}");
            }

            // Shell-specific rules
            if (domain.Contains("shell"))
            {
                revision.Priority = "High";
                // Shell requires extended documentation
                if (string.IsNullOrEmpty(revision.TechnicalChanges))
                {
                    revision.TechnicalChanges = "Detailed technical documentation required per Shell standards";
                }
            }
        }

        private static bool IsValidRevisionSequence(string currentRev, string newRev)
        {
            // Business rule: Validate revision progression (A→B, 01→02, etc.)
            try
            {
                // Handle alphabetic revisions (A, B, C...)
                if (currentRev.Length == 1 && newRev.Length == 1 &&
                    char.IsLetter(currentRev[0]) && char.IsLetter(newRev[0]))
                {
                    return newRev[0] > currentRev[0];
                }

                // Handle numeric revisions (01, 02, 03...)
                if (int.TryParse(currentRev, out int currentNum) && int.TryParse(newRev, out int newNum))
                {
                    return newNum > currentNum;
                }

                return true; // Allow other formats
            }
            catch
            {
                return true; // Default to allowing if validation fails
            }
        }

        private static string IncrementRevision(string currentRevision)
        {
            try
            {
                // Handle alphabetic revisions (A→B, B→C, etc.)
                if (currentRevision.Length == 1 && char.IsLetter(currentRevision[0]))
                {
                    return ((char)(currentRevision[0] + 1)).ToString();
                }

                // Handle numeric revisions (01→02, 02→03, etc.)
                if (int.TryParse(currentRevision, out int revNumber))
                {
                    return (revNumber + 1).ToString(currentRevision.Length == 2 ? "D2" : "D1");
                }

                // Default increment
                return "01";
            }
            catch
            {
                return "01";
            }
        }

        private static string DetermineRevisionComplexity(RevisionLine revision)
        {
            // Business rules for revision complexity
            if (IsMajorRevisionChange(revision.CurrentRevision, revision.NewRevision))
                return "Major";

            if (!string.IsNullOrEmpty(revision.DrawingNumber))
                return "Major";

            if (revision.RevisionReason?.ToLower().Contains("material") == true ||
                revision.RevisionReason?.ToLower().Contains("dimension") == true)
                return "Major";

            return "Minor";
        }

        private static bool IsMajorRevisionChange(string currentRev, string newRev)
        {
            // Business rule: Major revision if jumping multiple versions
            try
            {
                if (int.TryParse(currentRev, out int current) && int.TryParse(newRev, out int newVer))
                {
                    return (newVer - current) > 1;
                }

                if (currentRev.Length == 1 && newRev.Length == 1 &&
                    char.IsLetter(currentRev[0]) && char.IsLetter(newRev[0]))
                {
                    return (newRev[0] - currentRev[0]) > 1;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private static decimal CalculatePriceImpact(RevisionLine revision)
        {
            // Placeholder for price impact calculation
            // In real implementation, this would compare with existing pricing
            return 0.05m; // 5% default impact
        }

        private static void AssessTimelineImpact(RevisionLine revision)
        {
            // Business rule: Assess delivery timeline impact
            if (!string.IsNullOrEmpty(revision.RequestedDeliveryDate))
            {
                if (DateTime.TryParse(revision.RequestedDeliveryDate, out DateTime deliveryDate))
                {
                    var daysUntilDelivery = (deliveryDate - DateTime.Now).Days;
                    if (daysUntilDelivery < 14)
                    {
                        revision.Priority = "High";
                        processingLog.Add($"Rush delivery timeline detected for {revision.ArtiCode}");
                    }
                }
            }
        }

        private static string FormatRevisionNumber(string revision)
        {
            if (string.IsNullOrEmpty(revision)) return "00";

            revision = revision.Trim().ToUpper();

            // Ensure numeric revisions are padded (1 → 01, 2 → 02)
            if (int.TryParse(revision, out int num))
            {
                return num.ToString("D2");
            }

            return revision;
        }

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

        private static string CleanDescription(string description)
        {
            if (string.IsNullOrEmpty(description)) return "";

            description = Regex.Replace(description, @"\s+", " ").Trim();
            description = Regex.Replace(description, @"<[^>]*>", "");

            if (description.Length > 200)
                description = description.Substring(0, 197) + "...";

            return description;
        }

        private static string CleanRevisionReason(string reason)
        {
            if (string.IsNullOrEmpty(reason)) return "";

            reason = Regex.Replace(reason, @"\s+", " ").Trim();

            if (reason.Length > 150)
                reason = reason.Substring(0, 147) + "...";

            return reason;
        }

        #endregion

        #region Test Data Generation

        /// <summary>
        /// Generate test revisions for development/testing
        /// </summary>
        public static List<RevisionLine> GenerateTestRevisions()
        {
            var testRevisions = new List<RevisionLine>();

            try
            {
                // Major technical revision
                var technicalRevision = new RevisionLine();
                technicalRevision.ArtiCode = "123.456.789";
                technicalRevision.Description = "Pump assembly technical revision";
                technicalRevision.CurrentRevision = "A";
                technicalRevision.NewRevision = "B";
                technicalRevision.RevisionReason = "Material specification change from stainless steel to duplex steel";
                technicalRevision.TechnicalChanges = "Updated material specification, revised welding procedures";
                technicalRevision.CommercialChanges = "Price increase due to material upgrade";
                technicalRevision.DrawingNumber = "DWG-PUMP-001";
                technicalRevision.Priority = "High";
                technicalRevision.ApprovalRequired = "Yes";
                technicalRevision.RevisionStatus = "Draft";
                technicalRevision.SetExtractionDetails("TEST_GENERATION", "test.domain.com");

                testRevisions.Add(technicalRevision);

                // Minor drawing revision
                var drawingRevision = new RevisionLine();
                drawingRevision.ArtiCode = "234.567.890";
                drawingRevision.Description = "Valve assembly drawing correction";
                drawingRevision.CurrentRevision = "01";
                drawingRevision.NewRevision = "02";
                drawingRevision.RevisionReason = "Dimension tolerance correction";
                drawingRevision.TechnicalChanges = "Updated dimension tolerances on drawing";
                drawingRevision.CommercialChanges = "No commercial impact";
                drawingRevision.DrawingNumber = "DWG-VALVE-050";
                drawingRevision.Priority = "Normal";
                drawingRevision.ApprovalRequired = "Yes";
                drawingRevision.RevisionStatus = "Draft";
                drawingRevision.SetExtractionDetails("TEST_GENERATION", "test.domain.com");

                testRevisions.Add(drawingRevision);

                processingLog.Add($"Generated {testRevisions.Count} test revisions for development");
                return testRevisions;
            }
            catch (Exception ex)
            {
                processingLog.Add($"Error generating test revisions: {ex.Message}");
                return new List<RevisionLine>();
            }
        }

        #endregion
        public static async Task<List<RevisionLine>> ExtractRevisionsSafe(string emailBody, string emailDomain, string subject, AttachmentCollectionResponse attachments)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(emailBody) || string.IsNullOrWhiteSpace(emailDomain))
                {
                    return new List<RevisionLine>();
                }

                if (!emailDomain.StartsWith("@"))
                {
                    emailDomain = "@" + emailDomain;
                }

                // Use the existing IsRevisionEmail method with correct signature
                if (!RevisionLogicService.IsRevisionEmail(subject, emailBody, attachments))
                {
                    return new List<RevisionLine>();
                }

                // Use the existing ExtractRevisions method instead of ExtractRevisionsFromBody
                var revisions = RevisionLogicService.ExtractRevisions(emailBody, emailDomain, subject, attachments);

                return revisions;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in ExtractRevisionsSafe: {ex.Message}");
                return new List<RevisionLine>();
            }
        }
        // Add these methods to RevisionLogicService.cs to fix compilation errors

        #region Missing Method Implementations

        /// <summary>
        /// Check if email contains revision content
        /// </summary>
        #endregion
    }
}