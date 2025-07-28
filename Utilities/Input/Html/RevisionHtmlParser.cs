// ===============================================
// ENHANCED RevisionHtmlParser.cs - Better Revision Data Extraction
// Focuses on extracting real revision data from emails
// ===============================================

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Mkg_Elcotec_Automation.Models;

namespace Mkg_Elcotec_Automation.Utilities.Input.Html
{
    public class RevisionHtmlParser
    {
        public static HtmlDocument Doc { get; set; }
        private static List<string> DebugLog = new List<string>();

        public static void ClearDebugLog() => DebugLog.Clear();
        public static List<string> GetDebugLog() => new List<string>(DebugLog);

        private static void LogDebug(string message)
        {
            var logEntry = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
            DebugLog.Add(logEntry);
            Console.WriteLine($"[ENHANCED_REVISION_PARSER] {logEntry}");
        }

        /// <summary>
        /// 🔥 ENHANCED: Extract revision lines from HTML content with better detection
        /// </summary>
        public static List<RevisionLine> ExtractRevisionLines(string htmlContent, string clientDomain = "")
        {
            LogDebug("=== STARTING ENHANCED REVISION EXTRACTION ===");

            var revisionLines = new List<RevisionLine>();

            try
            {
                LogDebug($"Starting enhanced revision extraction (length: {htmlContent?.Length ?? 0})");

                if (string.IsNullOrEmpty(htmlContent))
                {
                    LogDebug("ERROR: HTML content is null or empty");
                    return revisionLines;
                }

                // Load HTML document
                var doc = new HtmlDocument();
                doc.LoadHtml(htmlContent);
                Doc = doc;

                // 🔥 STEP 1: Extract global revision info from email text
                var globalRevisionInfo = ExtractGlobalRevisionInfo(htmlContent);
                LogDebug($"Global revision info found: {globalRevisionInfo.revisionCount} revisions detected");

                // 🔥 STEP 2: Find revision indicators in the text
                var revisionIndicators = FindRevisionIndicators(htmlContent);
                LogDebug($"Found {revisionIndicators.Count} revision indicators");

                // 🔥 STEP 3: Find and process tables that might contain revisions
                var tables = doc.DocumentNode.SelectNodes("//table") ?? new HtmlNodeCollection(null);
                LogDebug($"Found {tables.Count} tables to analyze for revisions");

                foreach (var table in tables)
                {
                    var tableRevisions = ExtractRevisionsFromTable(table, globalRevisionInfo);
                    revisionLines.AddRange(tableRevisions);
                }

                // 🔥 STEP 4: If no table revisions found, create from text patterns
                if (!revisionLines.Any() && revisionIndicators.Any())
                {
                    var textRevisions = CreateRevisionsFromTextPatterns(revisionIndicators, globalRevisionInfo);
                    revisionLines.AddRange(textRevisions);
                }

                LogDebug($"\n=== ENHANCED REVISION EXTRACTION COMPLETE ===");
                LogDebug($"Successfully extracted {revisionLines.Count} enhanced revision lines");

                return revisionLines;
            }
            catch (Exception ex)
            {
                LogDebug($"CRITICAL ERROR in enhanced revision extraction: {ex.Message}");
                return revisionLines;
            }
        }

        /// <summary>
        /// 🔥 NEW: Extract global revision information from email text
        /// </summary>
        private static (int revisionCount, List<string> articleCodes, string changeDescription) ExtractGlobalRevisionInfo(string htmlContent)
        {
            try
            {
                LogDebug("🔍 Extracting global revision info from email content");

                // Convert HTML to plain text for analysis
                var doc = new HtmlDocument();
                doc.LoadHtml(htmlContent);
                var plainText = doc.DocumentNode.InnerText;

                // Find article codes that might have revisions
                var articleCodes = ExtractArticleCodesFromText(plainText);
                LogDebug($"Found {articleCodes.Count} potential article codes for revision");

                // Look for revision change indicators
                var changeDescription = ExtractChangeDescription(plainText);
                LogDebug($"Change description: {changeDescription}");

                // Count potential revisions
                var revisionCount = CountRevisionIndicators(plainText);
                LogDebug($"Estimated revision count: {revisionCount}");

                return (revisionCount, articleCodes, changeDescription);
            }
            catch (Exception ex)
            {
                LogDebug($"❌ Error extracting global revision info: {ex.Message}");
                return (0, new List<string>(), "");
            }
        }

        /// <summary>
        /// 🔥 Extract article codes from text
        /// </summary>
        private static List<string> ExtractArticleCodesFromText(string text)
        {
            var articleCodes = new List<string>();

            try
            {
                // Article code patterns
                var patterns = new[]
                {
                    @"(\d{3}\.\d{3}\.\d{3,4})",           // 891.029.1541
                    @"(\d{3}\.\d{2}\.\d{2})",             // 891.75.61
                    @"(\d{3}\.\d{3}\.\d{4}[A-Z]\d)",      // 891.75.01.000A1
                };

                foreach (var pattern in patterns)
                {
                    var matches = Regex.Matches(text, pattern);
                    foreach (Match match in matches)
                    {
                        var code = match.Groups[1].Value;
                        if (!articleCodes.Contains(code))
                        {
                            articleCodes.Add(code);
                            LogDebug($"📦 Found article code for revision: {code}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogDebug($"❌ Error extracting article codes: {ex.Message}");
            }

            return articleCodes;
        }

        /// <summary>
        /// 🔥 Extract change description from text
        /// </summary>
        private static string ExtractChangeDescription(string text)
        {
            try
            {
                // Look for common change description patterns
                var patterns = new[]
                {
                    @"(?:change|revision|update|modification)[:\s]+([^\.]+)",
                    @"(?:aanpassing|wijziging|verandering)[:\s]+([^\.]+)",
                    @"rev[:\s\.]+([A-Z0-9]+)[:\s]+([^\.]+)",
                };

                foreach (var pattern in patterns)
                {
                    var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        var description = match.Groups[match.Groups.Count - 1].Value.Trim();
                        if (description.Length > 10 && description.Length < 200)
                        {
                            LogDebug($"📝 Found change description: {description}");
                            return description;
                        }
                    }
                }

                return "Revision update";
            }
            catch
            {
                return "Revision update";
            }
        }

        /// <summary>
        /// 🔥 Count revision indicators in text
        /// </summary>
        private static int CountRevisionIndicators(string text)
        {
            try
            {
                var indicators = new[]
                {
                    @"rev[:\s\.]*[A-Z0-9]+",
                    @"revision[:\s]*[A-Z0-9]+",
                    @"A\d+→A\d+",
                    @"A\d+\s*to\s*A\d+",
                };

                int count = 0;
                foreach (var pattern in indicators)
                {
                    var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
                    count += matches.Count;
                }

                return Math.Max(1, count); // At least 1 if we're looking for revisions
            }
            catch
            {
                return 1;
            }
        }

        /// <summary>
        /// 🔥 Find specific revision indicators in text
        /// </summary>
        private static List<(string articleCode, string oldRevision, string newRevision, string description)> FindRevisionIndicators(string htmlContent)
        {
            var indicators = new List<(string, string, string, string)>();

            try
            {
                var doc = new HtmlDocument();
                doc.LoadHtml(htmlContent);
                var text = doc.DocumentNode.InnerText;

                // Pattern for revision changes: "891.75.01.000A1 + 891.75.02.000A1"
                var revisionPattern = @"(\d{3}\.\d{2}\.\d{2})\.000(A\d+)(?:\s*\+\s*(\d{3}\.\d{2}\.\d{2})\.000(A\d+))?";
                var matches = Regex.Matches(text, revisionPattern);

                foreach (Match match in matches)
                {
                    var articleCode = match.Groups[1].Value;
                    var newRevision = match.Groups[2].Value;
                    var oldRevision = "A0"; // Default old revision

                    // Try to determine old revision (usually one version back)
                    if (newRevision == "A1") oldRevision = "A0";
                    else if (newRevision == "A2") oldRevision = "A1";
                    else if (newRevision == "A3") oldRevision = "A2";

                    var description = $"Revision change from {oldRevision} to {newRevision}";

                    indicators.Add((articleCode, oldRevision, newRevision, description));
                    LogDebug($"🔄 Revision indicator: {articleCode} {oldRevision}→{newRevision}");

                    // If there's a second article code in the match
                    if (match.Groups.Count > 4 && !string.IsNullOrEmpty(match.Groups[3].Value))
                    {
                        var articleCode2 = match.Groups[3].Value;
                        var newRevision2 = match.Groups[4].Value;
                        var oldRevision2 = "A0";

                        if (newRevision2 == "A1") oldRevision2 = "A0";
                        else if (newRevision2 == "A2") oldRevision2 = "A1";
                        else if (newRevision2 == "A3") oldRevision2 = "A2";

                        var description2 = $"Revision change from {oldRevision2} to {newRevision2}";

                        indicators.Add((articleCode2, oldRevision2, newRevision2, description2));
                        LogDebug($"🔄 Revision indicator: {articleCode2} {oldRevision2}→{newRevision2}");
                    }
                }

                // Look for other revision patterns
                var simpleRevisionPattern = @"(\d{3}\.\d{3}\.\d{3,4})[:\s]*rev[:\s\.]*([A-Z0-9]+)";
                var simpleMatches = Regex.Matches(text, simpleRevisionPattern, RegexOptions.IgnoreCase);

                foreach (Match match in simpleMatches)
                {
                    var articleCode = match.Groups[1].Value;
                    var newRevision = match.Groups[2].Value;
                    var oldRevision = "A0";
                    var description = $"Revision update to {newRevision}";

                    // Avoid duplicates
                    if (!indicators.Any(i => i.Item1 == articleCode && i.Item3 == newRevision))
                    {
                        indicators.Add((articleCode, oldRevision, newRevision, description));
                        LogDebug($"🔄 Simple revision: {articleCode} →{newRevision}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogDebug($"❌ Error finding revision indicators: {ex.Message}");
            }

            return indicators;
        }

        /// <summary>
        /// 🔥 Extract revisions from HTML table
        /// </summary>
        private static List<RevisionLine> ExtractRevisionsFromTable(HtmlNode table, (int revisionCount, List<string> articleCodes, string changeDescription) globalInfo)
        {
            var revisions = new List<RevisionLine>();

            try
            {
                LogDebug("🔍 Analyzing table for revision data");

                var rows = table.SelectNodes(".//tr");
                if (rows == null) return revisions;

                int rowIndex = 0;
                foreach (var row in rows)
                {
                    rowIndex++;
                    var columns = row.SelectNodes(".//td");
                    if (columns == null || columns.Count < 2) continue;

                    // Look for revision indicators in any column
                    foreach (var column in columns)
                    {
                        var text = column.InnerText;

                        // Check if this cell contains revision information
                        if (ContainsRevisionInfo(text))
                        {
                            var revision = CreateRevisionFromTableCell(text, globalInfo.changeDescription, rowIndex);
                            if (revision != null)
                            {
                                revisions.Add(revision);
                                LogDebug($"Row {rowIndex}: ✅ Revision created from table cell");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogDebug($"❌ Error extracting revisions from table: {ex.Message}");
            }

            return revisions;
        }

        /// <summary>
        /// 🔥 Check if text contains revision information
        /// </summary>
        private static bool ContainsRevisionInfo(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;

            var revisionPatterns = new[]
            {
                @"\d{3}\.\d{2}\.\d{2}\.000A\d+",    // 891.75.01.000A1
                @"\d{3}\.\d{3}\.\d{3,4}",           // 891.029.1541
                @"rev[:\s\.]*[A-Z0-9]+",            // rev: A1
                @"A\d+→A\d+",                       // A0→A1
            };

            return revisionPatterns.Any(pattern => Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase));
        }

        /// <summary>
        /// 🔥 Create revision from table cell content
        /// </summary>
        private static RevisionLine CreateRevisionFromTableCell(string cellText, string changeDescription, int rowIndex)
        {
            try
            {
                // Extract article code
                var articleMatch = Regex.Match(cellText, @"(\d{3}\.\d{2}\.\d{2})(?:\.000(A\d+))?");
                if (!articleMatch.Success)
                {
                    articleMatch = Regex.Match(cellText, @"(\d{3}\.\d{3}\.\d{3,4})");
                }

                if (!articleMatch.Success) return null;

                var articleCode = articleMatch.Groups[1].Value;
                var newRevision = articleMatch.Groups.Count > 2 && !string.IsNullOrEmpty(articleMatch.Groups[2].Value)
                    ? articleMatch.Groups[2].Value : "A1";
                var oldRevision = "A0";

                // Try to determine old revision
                if (newRevision == "A1") oldRevision = "A0";
                else if (newRevision == "A2") oldRevision = "A1";
                else if (newRevision == "A3") oldRevision = "A2";

                // ✅ FIXED: Use correct constructor parameters
                var revision = new RevisionLine();
                revision.ArtiCode = articleCode;  // ✅ Set property directly
                revision.Description = string.IsNullOrEmpty(changeDescription) ? $"Revision update for {articleCode}" : changeDescription;
                revision.CurrentRevision = oldRevision;  // ✅ Use CurrentRevision instead of oldRevision
                revision.NewRevision = newRevision;
                revision.RevisionDate = DateTime.Now.ToString("dd-MM-yyyy");  // ✅ Use RevisionDate instead of changeDate
                revision.RevisionReason = "Email revision request";  // ✅ Use RevisionReason instead of changeReason
                // revision.requestedBy = "Customer";  // ❌ This property doesn't exist
                // revision.approvedBy = "";  // ❌ This property doesn't exist
                revision.RevisionStatus = "Pending";
                revision.Priority = "Normal";
                revision.ExtractionMethod = "ENHANCED_HTML_TABLE";
                revision.ExtractionDomain = "weir.com";

                LogDebug($"✅ Table revision created: {articleCode} {oldRevision}→{newRevision}");
                return revision;
            }
            catch (Exception ex)
            {
                LogDebug($"❌ Error creating revision from table cell: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 🔥 Create revisions from text patterns when no table data found
        /// </summary>
        /// </summary>
        private static List<RevisionLine> CreateRevisionsFromTextPatterns(
            List<(string articleCode, string oldRevision, string newRevision, string description)> indicators,
            (int revisionCount, List<string> articleCodes, string changeDescription) globalInfo)
        {
            var revisions = new List<RevisionLine>();

            try
            {
                LogDebug("🔍 Creating revisions from text patterns");

                foreach (var indicator in indicators)
                {
                    // ✅ FIXED: Use correct constructor parameters (no 'artiCode' named parameter)
                    var revision = new RevisionLine();
                    revision.ArtiCode = indicator.articleCode;  // ✅ Set property directly
                    revision.Description = !string.IsNullOrEmpty(globalInfo.changeDescription) ? globalInfo.changeDescription : indicator.description;
                    revision.CurrentRevision = indicator.oldRevision;  // ✅ Use CurrentRevision instead of oldRevision
                    revision.NewRevision = indicator.newRevision;
                    revision.RevisionDate = DateTime.Now.ToString("dd-MM-yyyy");
                    revision.RevisionReason = "Email revision request";  // ✅ Use RevisionReason instead of changeReason
                    // revision.requestedBy = "Customer";  // ❌ This property doesn't exist
                    // revision.approvedBy = "";  // ❌ This property doesn't exist  
                    revision.RevisionStatus = "Pending";
                    revision.Priority = "Normal";
                    revision.ExtractionMethod = "ENHANCED_TEXT_PATTERN";
                    revision.ExtractionDomain = "weir.com";

                    revisions.Add(revision);
                    LogDebug($"✅ Text pattern revision created: {indicator.articleCode} {indicator.oldRevision}→{indicator.newRevision}");
                }

                // If no specific indicators found but we have article codes, create generic revisions
                if (!revisions.Any() && globalInfo.articleCodes.Any())
                {
                    foreach (var articleCode in globalInfo.articleCodes.Take(3)) // Limit to avoid spam
                    {
                        // ✅ FIXED: Use correct constructor parameters
                        var revision = new RevisionLine();
                        revision.ArtiCode = articleCode;  // ✅ Set property directly
                        revision.Description = !string.IsNullOrEmpty(globalInfo.changeDescription) ? globalInfo.changeDescription : $"Revision update for {articleCode}";
                        revision.CurrentRevision = "A0";  // ✅ Use CurrentRevision instead of oldRevision
                        revision.NewRevision = "A1";
                        revision.RevisionDate = DateTime.Now.ToString("dd-MM-yyyy");
                        revision.RevisionReason = "Email revision request";  // ✅ Use RevisionReason instead of changeReason
                        // revision.requestedBy = "Customer";  // ❌ This property doesn't exist
                        // revision.approvedBy = "";  // ❌ This property doesn't exist
                        revision.RevisionStatus = "Pending";
                        revision.Priority = "Normal";
                        revision.ExtractionMethod = "ENHANCED_GENERIC";
                        revision.ExtractionDomain = "weir.com";

                        revisions.Add(revision);
                        LogDebug($"✅ Generic revision created: {articleCode} A0→A1");
                    }
                }
            }
            catch (Exception ex)
            {
                LogDebug($"❌ Error creating revisions from text patterns: {ex.Message}");
            }

            return revisions;
        }

        /// <summary>
        /// 🔥 Enhanced method specifically for Weir domain
        /// </summary>
        public static List<RevisionLine> ExtractRevisionLinesWeirDomain(string htmlContent)
        {
            LogDebug("=== STARTING ENHANCED WEIR DOMAIN REVISION EXTRACTION ===");

            try
            {
                // Use the enhanced method with Weir-specific optimizations
                var revisions = ExtractRevisionLines(htmlContent, "weir.com");

                // Apply Weir-specific post-processing
                foreach (var revision in revisions)
                {
                    // Ensure Weir domain is set
                    if (string.IsNullOrEmpty(revision.ExtractionDomain))
                    {
                        revision.ExtractionDomain = "weir.com";
                    }

                    // Set Weir-specific extraction method
                    if (revision.ExtractionMethod.Contains("ENHANCED"))
                    {
                        revision.ExtractionMethod = "ENHANCED_WEIR_" + revision.ExtractionMethod.Replace("ENHANCED_", "");
                    }
                }

                LogDebug($"✅ Weir-specific processing complete: {revisions.Count} revisions");
                return revisions;
            }
            catch (Exception ex)
            {
                LogDebug($"❌ Error in Weir domain revision extraction: {ex.Message}");
                return new List<RevisionLine>();
            }
        }

        // ===============================================
        // BACKWARD COMPATIBILITY METHODS
        // ===============================================

        public static List<RevisionLine> ExtractRevisionLinesOutlookForwardMail(string htmlContent)
        {
            LogDebug("=== FALLBACK: Using enhanced extraction for Outlook forward mail ===");
            return ExtractRevisionLines(htmlContent, "outlook.com");
        }
    }
}