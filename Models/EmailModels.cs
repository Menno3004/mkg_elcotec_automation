using System;
using System.Collections.Generic;

namespace Mkg_Elcotec_Automation.Models.EmailModels
{
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

    /// <summary>
    /// Email detail for workflow processing
    /// </summary>
    public class EmailDetail
    {
        public string Subject { get; set; }
        public string Sender { get; set; }
        public DateTime ReceivedDate { get; set; }
        public string ClientDomain { get; set; }
        public List<dynamic> Orders { get; set; } = new List<dynamic>();

        // ADD THIS:
        public List<AttachmentInfo> Attachments { get; set; } = new List<AttachmentInfo>();
        public string Body { get; internal set; }
        public int Quotes { get; internal set; }
        public int Revisions { get; internal set; }
        public string Domain { get; internal set; }
    }

    public class AttachmentInfo
    {
        public string Name { get; set; }
        public string Extension { get; set; }
        public long Size { get; set; }
        public string ContentType { get; set; }
    }

    /// <summary>
    /// Email import summary for workflow
    /// </summary>
    public class EmailImportSummary
    {
        public int TotalEmails { get; set; }
        public int ProcessedEmails { get; set; }
        public int FailedEmails { get; set; }
        public int TotalOrdersExtracted { get; set; }
        public int TotalOrderLinesExtracted { get; set; }
        public int TotalQuotesExtracted { get; set; }
        public int TotalQuoteLinesExtracted { get; set; }
        public int TotalRevisionsExtracted { get; set; }
        public List<EmailDetail> EmailDetails { get; set; } = new List<EmailDetail>();
        public List<SkippedEmailDetail> SkippedEmails { get; set; } = new List<SkippedEmailDetail>();

        // ✅ FIXED: Add separate list for duplicate emails
        public List<SkippedEmailDetail> DuplicateEmails { get; set; } = new List<SkippedEmailDetail>();

        public int SkippedEmailsCount => SkippedEmails?.Count ?? 0;
        // ✅ FIXED: Property for duplicate count
        public int DuplicateEmailsCount => DuplicateEmails?.Count ?? 0;

        // Keep existing properties
        public object ProcessingTime { get; set; }
        public int TotalRevisionLines { get; internal set; }
        public int TotalEmailsRetrieved { get; internal set; }
        public List<OrderLine> ExtractedOrderLines { get; internal set; }
        public List<QuoteLine> ExtractedQuoteLines { get; internal set; }
        public List<RevisionLine> ExtractedRevisionLines { get; internal set; }
        public int EmailDuplicatesRemoved { get; set; } = 0;

        public string GetOverallSummary()
        {
            return $"Processed {ProcessedEmails}/{TotalEmails} emails - " +
                   $"Orders: {TotalOrdersExtracted} ({TotalOrderLinesExtracted} lines), " +
                   $"Quotes: {TotalQuotesExtracted} ({TotalQuoteLinesExtracted} lines), " +
                   $"Revisions: {TotalRevisionsExtracted}";
        }
    }

    public class SkippedEmailDetail
    {
        public string Subject { get; set; }
        public string Reason { get; set; }
        public DateTime ProcessedAt { get; set; }
        public string Sender { get; set; }
        public string Domain { get; set; }
        public DateTime ReceivedDate { get; internal set; }
        public string ClientDomain { get; internal set; }
        public string SkipReason { get; internal set; }
    }

    /// <summary>
    /// MKG extraction results for workflow processing
    /// </summary>
    public class MkgExtractionResults
    {
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public TimeSpan ProcessingTime { get; set; }
        public bool Success { get; set; } = true;
        public string ErrorMessage { get; set; }

        public List<OrderLine> Orders { get; set; } = new List<OrderLine>();
        public List<QuoteLine> Quotes { get; set; } = new List<QuoteLine>();
        public List<RevisionLine> Revisions { get; set; } = new List<RevisionLine>();

        public int TotalExtracted => Orders.Count + Quotes.Count + Revisions.Count;

        public string GetSummary()
        {
            return $"Extracted {TotalExtracted} items: {Orders.Count} orders, {Quotes.Count} quotes, {Revisions.Count} revisions";
        }
    }
}