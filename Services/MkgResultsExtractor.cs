using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

public class MkgResultsExtractor
{
    public class MkgResultsData
    {
        public DateTime ExportTime { get; set; }
        public string User { get; set; }
        public MkgSummaryStats Summary { get; set; } = new MkgSummaryStats();
        public List<MkgOrderGroup> OrderGroups { get; set; } = new List<MkgOrderGroup>();
        public List<MkgQuoteGroup> QuoteGroups { get; set; } = new List<MkgQuoteGroup>();
        public List<MkgRevisionGroup> RevisionGroups { get; set; } = new List<MkgRevisionGroup>();
        public List<string> ProcessingSteps { get; set; } = new List<string>();
    }

    public class MkgSummaryStats
    {
        public int TotalItemsProcessed { get; set; }
        public int SuccessfullyInjected { get; set; }
        public int DuplicatesSkipped { get; set; }
        public int RealFailures { get; set; }
        public double EffectiveSuccessRate { get; set; }
        public int HeadersCreated { get; set; }
        public int LinesProcessed { get; set; }
        public string ProcessingTime { get; set; }
    }

    public class MkgOrderGroup
    {
        public string PoNumber { get; set; }
        public string MkgOrderId { get; set; }
        public int SuccessCount { get; set; }
        public int DuplicateCount { get; set; }
        public int FailureCount { get; set; }
        public List<MkgOrderItem> Items { get; set; } = new List<MkgOrderItem>();
    }

    public class MkgOrderItem
    {
        public string ArticleCode { get; set; }
        public string Status { get; set; }
        public string HttpStatus { get; set; }
        public DateTime ProcessedAt { get; set; }
        public string MkgOrderId { get; set; }
        public string ErrorMessage { get; set; }
        public bool IsSuccess => Status == "200";
        public bool IsDuplicate => Status == "DUPLICATE_SKIPPED";
        public string RequestPreview { get; set; }
        public string ResponsePreview { get; set; }
        public string Description { get; set; }
    }

    public class MkgQuoteGroup
    {
        public string RfqNumber { get; set; }
        public string MkgQuoteId { get; set; }
        public int LineCount { get; set; }
        public DateTime ProcessedAt { get; set; }
        public List<MkgQuoteItem> Items { get; set; } = new List<MkgQuoteItem>();
    }

    public class MkgQuoteItem
    {
        public string ArticleCode { get; set; }
        public string QuotedPrice { get; set; }
        public string Currency { get; set; }
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string MkgQuoteId { get; set; }
        public DateTime ProcessedAt { get; set; }
        public string HttpStatus { get; set; }
        public string RequestPreview { get; set; }
        public string ResponsePreview { get; set; }
    }

    public class MkgRevisionGroup
    {
        public string ArticleCode { get; set; }
        public string MkgRevisionId { get; set; }
        public string CurrentRevision { get; set; }
        public string NewRevision { get; set; }
        public DateTime ProcessedAt { get; set; }
        public List<MkgRevisionItem> Items { get; set; } = new List<MkgRevisionItem>();
    }

    public class MkgRevisionItem
    {
        public string ArticleCode { get; set; }
        public string CurrentRevision { get; set; }
        public string NewRevision { get; set; }
        public string ChangeReason { get; set; }
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string MkgRevisionId { get; set; }
        public DateTime ProcessedAt { get; set; }
        public string HttpStatus { get; set; }
        public string RequestPreview { get; set; }
        public string ResponsePreview { get; set; }
    }

    public static MkgResultsData ExtractMkgResults(string resultsText)
    {
        if (string.IsNullOrEmpty(resultsText))
            return new MkgResultsData();

        var data = new MkgResultsData();
        var lines = resultsText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

        try
        {
            // Extract header information
            ExtractHeaderInfo(lines, data);

            // Extract summary statistics
            ExtractSummaryStats(lines, data);

            // Extract order details
            ExtractOrderDetails(lines, data);

            // Extract quote details
            ExtractQuoteDetails(lines, data);

            // Extract revision details
            ExtractRevisionDetails(lines, data);

            // Extract processing steps
            ExtractProcessingSteps(lines, data);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error extracting MKG results: {ex.Message}");
        }

        return data;
    }

    private static void ExtractHeaderInfo(string[] lines, MkgResultsData data)
    {
        foreach (var line in lines)
        {
            if (line.Contains("Exported at:") || line.Contains("Generated:"))
            {
                var parts = line.Split(new[] { ": " }, StringSplitOptions.None);
                var timeStr = parts.Length > 1 ? parts[parts.Length - 1] : "";
                if (DateTime.TryParse(timeStr, out var exportTime))
                    data.ExportTime = exportTime;
            }
            else if (line.Contains("User:"))
            {
                var parts = line.Split(new[] { ": " }, StringSplitOptions.None);
                data.User = parts.Length > 1 ? parts[parts.Length - 1].Trim() : "";
            }
        }
    }

    private static void ExtractSummaryStats(string[] lines, MkgResultsData data)
    {
        foreach (var line in lines)
        {
            if (line.Contains("Total Items Processed:"))
            {
                var totalItems = 0;
                int.TryParse(ExtractNumber(line), out totalItems);
                data.Summary.TotalItemsProcessed = totalItems;
            }
            else if (line.Contains("Successfully Injected:"))
            {
                var successfulInjections = 0;
                int.TryParse(ExtractNumber(line), out successfulInjections);
                data.Summary.SuccessfullyInjected = successfulInjections;
            }
            else if (line.Contains("Duplicates Skipped:"))
            {
                var duplicatesSkipped = 0;
                int.TryParse(ExtractNumber(line), out duplicatesSkipped);
                data.Summary.DuplicatesSkipped = duplicatesSkipped;
            }
            else if (line.Contains("Real Failures:"))
            {
                var realFailures = 0;
                int.TryParse(ExtractNumber(line), out realFailures);
                data.Summary.RealFailures = realFailures;
            }
            else if (line.Contains("Effective Success Rate:"))
            {
                var parts = line.Split(':');
                var rateStr = parts.Length > 1 ? parts[parts.Length - 1].Replace("%", "").Trim() : "0";
                if (double.TryParse(rateStr.Replace(",", "."), out var rate))
                    data.Summary.EffectiveSuccessRate = rate;
            }
            else if (line.Contains("Headers Created:"))
            {
                var headersCreated = 0;
                int.TryParse(ExtractNumber(line), out headersCreated);
                data.Summary.HeadersCreated = headersCreated;
            }
            else if (line.Contains("Lines Processed:"))
            {
                var linesProcessed = 0;
                int.TryParse(ExtractNumber(line), out linesProcessed);
                data.Summary.LinesProcessed = linesProcessed;
            }
            else if (line.Contains("Processing Time:"))
            {
                var parts = line.Split(':');
                data.Summary.ProcessingTime = parts.Length > 1 ? parts[parts.Length - 1].Trim() : "";
            }
        }
    }

    private static void ExtractOrderDetails(string[] lines, MkgResultsData data)
    {
        var currentGroup = (MkgOrderGroup)null;
        var currentItem = (MkgOrderItem)null;
        var inOrderSection = false;

        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i];

            if (line.Contains("ORDER INJECTION RESULTS") || line.Contains("Order Details"))
            {
                inOrderSection = true;
                continue;
            }

            if (line.Contains("QUOTE INJECTION RESULTS") || line.Contains("REVISION INJECTION RESULTS"))
            {
                inOrderSection = false;
                continue;
            }

            if (!inOrderSection) continue;

            // Parse PO groups
            if (line.Contains("PO:") && line.Contains("(✅") && line.Contains("🔄") && line.Contains("❌"))
            {
                if (currentGroup != null)
                    data.OrderGroups.Add(currentGroup);

                currentGroup = new MkgOrderGroup();

                // Extract PO number
                var poMatch = System.Text.RegularExpressions.Regex.Match(line, @"PO:\s*([^\s]+)");
                if (poMatch.Success)
                    currentGroup.PoNumber = poMatch.Groups[1].Value;

                // Extract counts
                var successMatch = Regex.Match(line, @"✅(\d+)");
                if (successMatch.Success)
                {
                    var successCount = 0;
                    int.TryParse(successMatch.Groups[1].Value, out successCount);
                    currentGroup.SuccessCount = successCount;
                }

                var duplicateMatch = Regex.Match(line, @"🔄(\d+)");
                if (duplicateMatch.Success)
                {
                    var duplicateCount = 0;
                    int.TryParse(duplicateMatch.Groups[1].Value, out duplicateCount);
                    currentGroup.DuplicateCount = duplicateCount;
                }

                var failMatch = Regex.Match(line, @"❌(\d+)");
                if (failMatch.Success)
                {
                    var failCount = 0;
                    int.TryParse(failMatch.Groups[1].Value, out failCount);
                    currentGroup.FailureCount = failCount;
                }
            }
            // Parse individual items
            else if ((line.Contains("✅") || line.Contains("🔄") || line.Contains("❌")) && line.Contains("Status:"))
            {
                currentItem = new MkgOrderItem();

                // Extract article code
                var parts = line.Split('|');
                if (parts.Length > 0)
                {
                    var articlePart = parts[0].Trim();
                    var articleMatch = Regex.Match(articlePart, @"[✅🔄❌]\s*(.+)");
                    if (articleMatch.Success)
                        currentItem.ArticleCode = articleMatch.Groups[1].Value.Trim();
                }

                // Extract status
                var statusMatch = Regex.Match(line, @"Status:([^|]+)");
                if (statusMatch.Success)
                    currentItem.HttpStatus = statusMatch.Groups[1].Value.Trim();

                // Extract time
                var timeMatch = Regex.Match(line, @"Time:(\d{2}:\d{2}:\d{2})");
                if (timeMatch.Success && DateTime.TryParse($"2025-07-27 {timeMatch.Groups[1].Value}", out var time))
                    currentItem.ProcessedAt = time;

                currentItem.Status = currentItem.HttpStatus;
                currentGroup?.Items.Add(currentItem);
            }
            // Extract nested details
            else if (currentItem != null && line.Contains("MKG Order ID:"))
            {
                var parts = line.Split(':');
                currentItem.MkgOrderId = parts.Length > 1 ? parts[parts.Length - 1].Trim() : "";
                if (currentGroup != null && string.IsNullOrEmpty(currentGroup.MkgOrderId))
                    currentGroup.MkgOrderId = currentItem.MkgOrderId;
            }
            else if (currentItem != null && line.Contains("Description:"))
            {
                var parts = line.Split(':');
                currentItem.Description = parts.Length > 1 ? parts[parts.Length - 1].Trim() : "";
            }
            else if (currentItem != null && line.Contains("Error:"))
            {
                var parts = line.Split(':');
                currentItem.ErrorMessage = parts.Length > 1 ? parts[parts.Length - 1].Trim() : "";
            }
        }

        if (currentGroup != null)
            data.OrderGroups.Add(currentGroup);
    }

    private static void ExtractQuoteDetails(string[] lines, MkgResultsData data)
    {
        var currentGroup = (MkgQuoteGroup)null;
        var inQuoteSection = false;

        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i];

            if (line.Contains("QUOTE INJECTION RESULTS") || line.Contains("Quote Headers"))
            {
                inQuoteSection = true;
                continue;
            }

            if (line.Contains("REVISION INJECTION RESULTS") || line.Contains("END OF QUOTE"))
            {
                inQuoteSection = false;
                continue;
            }

            if (!inQuoteSection) continue;

            // Parse quote headers
            if (line.Contains("PARENT: Quote Header"))
            {
                if (currentGroup != null)
                    data.QuoteGroups.Add(currentGroup);

                currentGroup = new MkgQuoteGroup();

                var headerMatch = Regex.Match(line, @"Quote Header\s+([^\s]+)");
                if (headerMatch.Success)
                    currentGroup.MkgQuoteId = headerMatch.Groups[1].Value;
            }
            else if (currentGroup != null && line.Contains("RFQ:"))
            {
                var rfqMatch = Regex.Match(line, @"RFQ:\s*([^|]+)");
                if (rfqMatch.Success)
                    currentGroup.RfqNumber = rfqMatch.Groups[1].Value.Trim();

                var linesMatch = Regex.Match(line, @"Lines:\s*(\d+)");
                if (linesMatch.Success)
                {
                    var lineCount = 0;
                    int.TryParse(linesMatch.Groups[1].Value, out lineCount);
                    currentGroup.LineCount = lineCount;
                }

                var timeMatch = Regex.Match(line, @"Time:\s*(\d{2}:\d{2}:\d{2})");
                if (timeMatch.Success && DateTime.TryParse($"2025-07-27 {timeMatch.Groups[1].Value}", out var time))
                    currentGroup.ProcessedAt = time;
            }
            else if (currentGroup != null && line.Contains("CHILD:") && line.Contains("→ ✓"))
            {
                var item = new MkgQuoteItem
                {
                    MkgQuoteId = currentGroup.MkgQuoteId,
                    ProcessedAt = currentGroup.ProcessedAt,
                    Success = true,
                    HttpStatus = "200"
                };

                var childMatch = Regex.Match(line, @"CHILD:\s*([^(]+)");
                if (childMatch.Success)
                    item.ArticleCode = childMatch.Groups[1].Value.Trim();

                var currencyMatch = Regex.Match(line, @"\(([^)]+)\)");
                if (currencyMatch.Success)
                    item.Currency = currencyMatch.Groups[1].Value;

                currentGroup.Items.Add(item);
            }
        }

        if (currentGroup != null)
            data.QuoteGroups.Add(currentGroup);
    }

    private static void ExtractRevisionDetails(string[] lines, MkgResultsData data)
    {
        // Similar pattern for revisions - implement when revision data is available
        var inRevisionSection = false;

        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i];

            if (line.Contains("REVISION INJECTION RESULTS"))
            {
                inRevisionSection = true;
                continue;
            }

            if (line.Contains("END OF REVISION") || line.Contains("INCREMENTAL PROCESSING"))
            {
                inRevisionSection = false;
                continue;
            }

            if (!inRevisionSection) continue;

            // Parse revision data when available
            // TODO: Implement revision parsing when revision output format is defined
        }
    }

    private static void ExtractProcessingSteps(string[] lines, MkgResultsData data)
    {
        var inStepsSection = false;

        foreach (var line in lines)
        {
            if (line.Contains("INCREMENTAL PROCESSING STEPS"))
            {
                inStepsSection = true;
                continue;
            }

            if (inStepsSection && line.Contains("===") && line.Contains("END"))
            {
                inStepsSection = false;
                continue;
            }

            if (inStepsSection && line.Trim().StartsWith("["))
            {
                data.ProcessingSteps.Add(line.Trim());
            }
        }
    }

    private static string ExtractNumber(string line)
    {
        var match = Regex.Match(line, @"(\d+)");
        return match.Success ? match.Groups[1].Value : "0";
    }

}
