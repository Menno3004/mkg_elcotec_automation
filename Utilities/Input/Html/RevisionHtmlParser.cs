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
        private static List<string> DebugLog = new List<string>();

        public static void ClearDebugLog() => DebugLog.Clear();
        public static List<string> GetDebugLog() => new List<string>(DebugLog);

        private static void LogDebug(string message)
        {
            var logEntry = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
            DebugLog.Add(logEntry);
            Console.WriteLine($"[REVISION_HTML_PARSER] {logEntry}");
        }

        /// <summary>
        /// Extract revision lines from HTML content
        /// </summary>
        public static List<RevisionLine> ExtractRevisionLines(string htmlContent, string clientDomain = "")
        {
            var revisionLines = new List<RevisionLine>();

            try
            {
                LogDebug($"Starting revision extraction from HTML content (length: {htmlContent?.Length ?? 0})");

                if (string.IsNullOrWhiteSpace(htmlContent))
                {
                    LogDebug("HTML content is null or empty");
                    return revisionLines;
                }

                // Load HTML document
                var doc = new HtmlDocument();
                doc.LoadHtml(htmlContent);

                // Try different extraction methods
                var extractedRevisions = new List<RevisionLine>();

                // Method 1: Look for revision tables
                extractedRevisions.AddRange(ExtractFromRevisionTables(doc, clientDomain));

                // Method 2: Look for revision patterns in text
                extractedRevisions.AddRange(ExtractFromTextPatterns(htmlContent, clientDomain));

                // Method 3: Look for ECN/ECR patterns
                extractedRevisions.AddRange(ExtractFromEcnPatterns(htmlContent, clientDomain));

                // Remove duplicates
                revisionLines = RemoveDuplicateRevisions(extractedRevisions);

                LogDebug($"✅ Extracted {revisionLines.Count} revision lines");
                return revisionLines;
            }
            catch (Exception ex)
            {
                LogDebug($"❌ Error extracting revision lines: {ex.Message}");
                return revisionLines;
            }
        }

        private static List<RevisionLine> ExtractFromRevisionTables(HtmlDocument doc, string clientDomain)
        {
            var revisions = new List<RevisionLine>();

            try
            {
                // Look for tables with revision-related content
                var revisionTables = doc.DocumentNode.SelectNodes("//table[contains(@class, 'revision') or contains(@id, 'revision') or contains(@class, 'ecn') or contains(@id, 'ecn')]");

                if (revisionTables != null)
                {
                    LogDebug($"Found {revisionTables.Count} potential revision tables");

                    foreach (var table in revisionTables)
                    {
                        var rows = table.SelectNodes(".//tr");
                        if (rows == null) continue;

                        LogDebug($"Processing revision table with {rows.Count} rows");

                        foreach (var row in rows.Skip(1)) // Skip header row
                        {
                            var cells = row.SelectNodes(".//td");
                            if (cells == null || cells.Count < 4) continue;

                            try
                            {
                                var revision = new RevisionLine
                                {
                                    LineNumber = GetNextLineNumber(revisions),
                                    ArtiCode = ExtractArtiCode(cells[0]?.InnerText?.Trim()),
                                    Description = cells[1]?.InnerText?.Trim() ?? "",
                                    CurrentRevision = cells[2]?.InnerText?.Trim() ?? "",
                                    NewRevision = cells[3]?.InnerText?.Trim() ?? "",
                                    RevisionReason = cells.Count > 4 ? cells[4]?.InnerText?.Trim() ?? "" : "",
                                    ClientDomain = clientDomain,
                                    ExtractionMethod = "REVISION_TABLE_EXTRACTION",
                                    ExtractionTimestamp = DateTime.Now
                                };

                                if (IsValidRevision(revision))
                                {
                                    revisions.Add(revision);
                                    LogDebug($"✅ Extracted revision: {revision.ArtiCode} ({revision.CurrentRevision} → {revision.NewRevision})");
                                }
                            }
                            catch (Exception ex)
                            {
                                LogDebug($"❌ Error processing revision table row: {ex.Message}");
                            }
                        }
                    }
                }
                else
                {
                    LogDebug("No revision tables found");
                }
            }
            catch (Exception ex)
            {
                LogDebug($"❌ Error extracting from revision tables: {ex.Message}");
            }

            return revisions;
        }

        private static List<RevisionLine> ExtractFromTextPatterns(string htmlContent, string clientDomain)
        {
            var revisions = new List<RevisionLine>();

            try
            {
                LogDebug("Extracting revisions from text patterns");

                // Pattern for revision changes: "Part ABC123 from Rev A to Rev B"
                var revisionPattern = @"(?i)(?:part|article|item)\s+([A-Z0-9\-\.]+)\s+(?:from|changed|revised)\s+(?:rev|revision)\s+([A-Z0-9]+)\s+(?:to|→)\s+(?:rev|revision)\s+([A-Z0-9]+)";
                var matches = Regex.Matches(htmlContent, revisionPattern);

                LogDebug($"Found {matches.Count} revision pattern matches");

                foreach (Match match in matches)
                {
                    if (match.Groups.Count >= 4)
                    {
                        var revision = new RevisionLine
                        {
                            LineNumber = GetNextLineNumber(revisions),
                            ArtiCode = match.Groups[1].Value.Trim(),
                            Description = "Revision change detected from email content",
                            CurrentRevision = match.Groups[2].Value.Trim(),
                            NewRevision = match.Groups[3].Value.Trim(),
                            RevisionReason = "Email notification",
                            ClientDomain = clientDomain,
                            ExtractionMethod = "TEXT_PATTERN_EXTRACTION",
                            ExtractionTimestamp = DateTime.Now
                        };

                        if (IsValidRevision(revision))
                        {
                            revisions.Add(revision);
                            LogDebug($"✅ Extracted revision from text: {revision.ArtiCode} ({revision.CurrentRevision} → {revision.NewRevision})");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogDebug($"❌ Error extracting from text patterns: {ex.Message}");
            }

            return revisions;
        }

        private static List<RevisionLine> ExtractFromEcnPatterns(string htmlContent, string clientDomain)
        {
            var revisions = new List<RevisionLine>();

            try
            {
                LogDebug("Extracting revisions from ECN/ECR patterns");

                // Pattern for ECN/ECR numbers with part numbers
                var ecnPattern = @"(?i)(?:ECN|ECR|ECO)[-\s]*([0-9]+).*?(?:part|article|item)\s+([A-Z0-9\-\.]+)";
                var matches = Regex.Matches(htmlContent, ecnPattern);

                LogDebug($"Found {matches.Count} ECN/ECR pattern matches");

                foreach (Match match in matches)
                {
                    if (match.Groups.Count >= 3)
                    {
                        var revision = new RevisionLine
                        {
                            LineNumber = GetNextLineNumber(revisions),
                            ArtiCode = match.Groups[2].Value.Trim(),
                            Description = $"ECN/ECR change - {match.Groups[1].Value}",
                            CurrentRevision = "A0", // Default assumption
                            NewRevision = "A1", // Default assumption
                            RevisionReason = $"ECN/ECR {match.Groups[1].Value}",
                            ClientDomain = clientDomain,
                            ExtractionMethod = "ECN_PATTERN_EXTRACTION",
                            ExtractionTimestamp = DateTime.Now
                        };

                        if (IsValidRevision(revision))
                        {
                            revisions.Add(revision);
                            LogDebug($"✅ Extracted revision from ECN: {revision.ArtiCode} (ECN: {match.Groups[1].Value})");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogDebug($"❌ Error extracting from ECN patterns: {ex.Message}");
            }

            return revisions;
        }

        private static List<RevisionLine> RemoveDuplicateRevisions(List<RevisionLine> revisions)
        {
            var uniqueRevisions = new List<RevisionLine>();
            var seenRevisions = new HashSet<string>();

            foreach (var revision in revisions)
            {
                var key = $"{revision.ArtiCode}-{revision.CurrentRevision}-{revision.NewRevision}";
                if (!seenRevisions.Contains(key))
                {
                    seenRevisions.Add(key);
                    uniqueRevisions.Add(revision);
                }
                else
                {
                    LogDebug($"⚠️ Duplicate revision filtered: {key}");
                }
            }

            LogDebug($"Removed {revisions.Count - uniqueRevisions.Count} duplicate revisions");
            return uniqueRevisions;
        }

        private static string ExtractArtiCode(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";

            // Look for common article code patterns
            var patterns = new[]
            {
                @"[A-Z]{2,3}[0-9]{3,6}",  // ABC123456
                @"[0-9]{6,8}",            // 12345678
                @"[A-Z0-9\-\.]{6,15}"     // General pattern
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern);
                if (match.Success)
                {
                    return match.Value;
                }
            }

            // If no pattern matches, return the cleaned text
            return Regex.Replace(text, @"[^A-Z0-9\-\.]", "");
        }

        private static bool IsValidRevision(RevisionLine revision)
        {
            return !string.IsNullOrEmpty(revision.ArtiCode) &&
                   !string.IsNullOrEmpty(revision.CurrentRevision) &&
                   !string.IsNullOrEmpty(revision.NewRevision) &&
                   revision.CurrentRevision != revision.NewRevision;
        }

        private static string GetNextLineNumber(List<RevisionLine> existingRevisions)
        {
            return (existingRevisions.Count + 1).ToString("D3");
        }
    }
}