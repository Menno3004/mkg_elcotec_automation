using System;
using System.Collections.Generic;

namespace Mkg_Elcotec_Automation.Models
{
    public enum TabStatus
    {
        Green,  // No errors, no duplicates
        Yellow, // Duplicates detected, no injection errors
        Red     // Injection errors present
    }

    public enum TabColorStatus
    {
        Green,      // No failures, all successful
        Orange,     // Warnings/partial failures (duplicates detected but handled)  
        Red         // Critical failures
    }
    /// <summary>
    /// Progress tracking for workflows
    /// </summary>
    public class WorkflowProgress
    {
        public int CurrentStep { get; set; }
        public int TotalSteps { get; set; }
        public string CurrentOperation { get; set; } = "";
        public string StatusMessage { get; set; } = "";
        public DateTime StartTime { get; set; } = DateTime.Now;
        public List<string> Log { get; set; } = new List<string>();

        public int PercentageComplete => TotalSteps > 0 ? (CurrentStep * 100) / TotalSteps : 0;

        public void UpdateProgress(int step, string operation, string message = "")
        {
            CurrentStep = step;
            CurrentOperation = operation;
            StatusMessage = message;
            Log.Add($"[{DateTime.Now:HH:mm:ss}] Step {step}/{TotalSteps}: {operation} - {message}");
        }
    }
    public class AutomationRunData
    {
        public RunHistoryItem RunInfo { get; set; }
        public List<object> EmailResults { get; set; } = new();
        public List<object> MkgResults { get; set; } = new();
        public List<object> ErrorResults { get; set; } = new();
        public Dictionary<string, object> Settings { get; set; } = new();

        // Existing injection failure tracking
        public bool HasInjectionFailures { get; set; } = false;
        public int TotalFailuresAtCompletion { get; set; } = 0;
        public DateTime? LastInjectionStatusUpdate { get; set; }

        // 🔥 NEW: Duplicate detection support
        public bool HasMkgDuplicatesDetected { get; set; } = false;
        public int TotalDuplicatesDetected { get; set; } = 0;
        public DateTime? LastDuplicateDetectionUpdate { get; set; }
        public bool HasInjectionWarnings { get; set; } = false; // NEW
        public TabColorStatus TabColorStatus { get; set; } = TabColorStatus.Green; // NEW
        public int TotalWarningsAtCompletion { get; set; } = 0; // NEW
    }

    public class ImportStatistics
    {
        public int EmailsProcessed { get; set; }
        public int OrdersFound { get; set; }
        public int QuotesFound { get; set; }
        public int RevisionsFound { get; set; }
        public int ErrorsEncountered { get; set; }
        public DateTime StartTime { get; set; } = DateTime.Now;
        public DateTime LastUpdateTime { get; set; } = DateTime.Now;
        public TimeSpan ElapsedTime => DateTime.Now - StartTime;

        public void UpdateStatistics(int emails, int orders, int quotes, int revisions, int errors)
        {
            EmailsProcessed = emails;
            OrdersFound = orders;
            QuotesFound = quotes;
            RevisionsFound = revisions;
            ErrorsEncountered = errors;
            LastUpdateTime = DateTime.Now;
        }

        public string GetSummaryText()
        {
            return $"📧 Emails: {EmailsProcessed} | 📦 Orders: {OrdersFound} | 💰 Quotes: {QuotesFound} | 🔄 Revisions: {RevisionsFound} | ❌ Errors: {ErrorsEncountered}";
        }
    }

    /// <summary>
    /// Validation result for import data
    /// </summary>
    public class ValidationResult
    {
        public bool IsValid { get; set; } = true;
        public List<string> Errors { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();
        public string ValidatedItemId { get; set; } = "";
        public string ValidatedItemType { get; set; } = "";

        public void AddError(string error)
        {
            Errors.Add(error);
            IsValid = false;
        }

        public void AddWarning(string warning)
        {
            Warnings.Add(warning);
        }

        public string GetSummary()
        {
            return IsValid ? "Valid" : $"Invalid: {Errors.Count} errors, {Warnings.Count} warnings";
        }
    }

    /// <summary>
    /// Import error information
    /// </summary>
    public class ImportError
    {
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public string ErrorType { get; set; } = "";
        public string ErrorMessage { get; set; } = "";
        public string EmailSubject { get; set; } = "";
        public string EmailSender { get; set; } = "";
        public string ExtractionMethod { get; set; } = "";
        public string StackTrace { get; set; } = "";
        public bool IsCritical { get; set; }

        public string GetFormattedError()
        {
            return $"[{Timestamp:HH:mm:ss}] {ErrorType}: {ErrorMessage} (Email: {EmailSubject})";
        }
    }

    /// <summary>
    /// Customer information
    /// </summary>
    public class CustomerInfo
    {
        public string CustomerName { get; set; } = "";
        public string CustomerCode { get; set; } = "";
        public string Domain { get; set; } = "";
        public string CustomerType { get; set; } = "STANDARD";
        public string ContactEmail { get; set; } = "";
        public string ContactName { get; set; } = "";
        public bool IsActive { get; set; } = true;
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        public string Notes { get; set; } = "";
        public DateTime CachedAt { get; set; } = DateTime.Now;

        // MKG properties with full names
        public string AdministrationNumber { get; set; } = "1";
        public string DebtorNumber { get; set; } = "30010";
        public string RelationNumber { get; set; } = "2";
        public string EmailDomain { get; set; } = "";
        public bool IsHighPriority { get; set; } = false;
    }
    /// <summary>
    /// Domain configuration for extraction
    /// </summary>
    public class ExtractionConfiguration
    {
        public string Domain { get; set; } = "";
        public string CustomerType { get; set; } = "STANDARD";
        public List<string> OrderIndicators { get; set; } = new List<string>();
        public List<string> QuoteIndicators { get; set; } = new List<string>();
        public List<string> RevisionIndicators { get; set; } = new List<string>();
        public List<string> ArticleCodePatterns { get; set; } = new List<string>();
        public bool EnableHtmlParsing { get; set; } = true;
        public bool EnableSubjectParsing { get; set; } = true;
        public bool EnableAttachmentParsing { get; set; } = true;
    }
}
