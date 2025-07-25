using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Mkg_Elcotec_Automation.Models
{
    #region Core API Response Models

    /// <summary>
    /// Generic MKG API response model
    /// </summary>
    public class MkgApiResponse
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; } = "";
        public string StatusCode { get; set; } = "";
        public string ResponseData { get; set; } = "";

        // Order-specific properties
        public string OrderNumber { get; set; } = "";
        public string OrderId { get; set; } = "";

        // Quote-specific properties
        public string QuoteNumber { get; set; } = "";
        public string QuoteId { get; set; } = "";

        // Revision-specific properties
        public string RevisionNumber { get; set; } = "";
        public string RevisionId { get; set; } = "";

        // Generic line item properties
        public string LineId { get; set; } = "";

        // Validation and error details
        public List<string> ValidationErrors { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();

        // Metadata
        public DateTime ResponseTime { get; set; } = DateTime.Now;
        public string ApiEndpoint { get; set; } = "";
        public string RequestId { get; set; } = "";
        public TimeSpan ProcessingTime { get; set; }

        // Helper methods
        public bool HasValidationErrors => ValidationErrors?.Count > 0;
        public bool HasWarnings => Warnings?.Count > 0;

        public void AddValidationError(string error)
        {
            if (ValidationErrors == null) ValidationErrors = new List<string>();
            ValidationErrors.Add(error);
        }

        public void AddWarning(string warning)
        {
            if (Warnings == null) Warnings = new List<string>();
            Warnings.Add(warning);
        }
    }

    #endregion

    #region Order Models

    /// <summary>
    /// Result of MKG order header creation
    /// </summary>
    public class MkgOrderHeaderResult
    {
        public bool Success { get; set; }
        public string MkgOrderNumber { get; set; }
        public string ErrorMessage { get; set; }
        public string HttpStatusCode { get; set; }
        public string RequestPayload { get; set; }
        public string ResponsePayload { get; set; }
    }

    /// <summary>
    /// Result of MKG order line injection
    /// </summary>
    public class MkgOrderLineResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string HttpStatusCode { get; set; }
        public string RequestPayload { get; set; }
        public string ResponsePayload { get; set; }
    }

    /// <summary>
    /// MKG Order injection result
    /// </summary>
    public class MkgOrderResult
    {
        public string LineId { get; set; }
        public string OrderLineId { get; set; }
        public string ArtiCode { get; set; }
        public string PoNumber { get; set; }
        public string Description { get; set; }
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public DateTime ProcessedAt { get; set; }
        public string HttpStatusCode { get; set; }
        public string StatusCode { get; set; }
        public string MkgOrderId { get; set; }
        public string OrderNumber { get; set; }
        public string ResponseData { get; set; }
        public List<string> ValidationErrors { get; set; } = new List<string>();
        public string RequestPayload { get; set; }
        public string ResponsePayload { get; set; }
        public bool IsDuplicate { get; set; } = false;
    }

    /// <summary>
    /// MKG Order injection summary
    /// </summary>
    public class MkgOrderInjectionSummary
    {
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public TimeSpan ProcessingTime { get; set; }
        public int TotalOrders { get; set; }
        public int SuccessfulInjections { get; set; }
        public int FailedInjections { get; set; }
        public int DuplicatesFiltered { get; set; }
        public int DuplicatesDetected { get; set; } = 0;
        public List<MkgOrderResult> OrderResults { get; set; } = new List<MkgOrderResult>();
        public List<string> Errors { get; set; } = new List<string>();
        public object DataQualityIssuesFixed { get; set; }
    }

    #endregion

    #region Quote Models

    /// <summary>
    /// Result of MKG quote header creation
    /// </summary>
    public class MkgQuoteHeaderResult
    {
        public bool Success { get; set; }
        public string MkgQuoteNumber { get; set; }
        public string ErrorMessage { get; set; }
        public string HttpStatusCode { get; set; }
        public string RequestPayload { get; set; }
        public string ResponsePayload { get; set; }
    }

    /// <summary>
    /// Result of MKG quote line injection
    /// </summary>
    public class MkgQuoteLineResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string HttpStatusCode { get; set; }
        public string RequestPayload { get; set; }
        public string ResponsePayload { get; set; }
    }

    /// <summary>
    /// MKG Quote injection result
    /// </summary>
    public class MkgQuoteResult
    {
        public string QuoteLineId { get; set; } = "";
        public string ArtiCode { get; set; }
        public string RfqNumber { get; set; }
        public string QuotedPrice { get; set; }
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string MkgQuoteId { get; set; }
        public DateTime ProcessedAt { get; set; }
        public string HttpStatusCode { get; set; }
        public List<string> ValidationErrors { get; set; } = new List<string>();
        public string RequestPayload { get; set; }
        public string ResponsePayload { get; set; }
    }

    /// <summary>
    /// MKG Quote injection summary
    /// </summary>
    public class MkgQuoteInjectionSummary
    {
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public TimeSpan ProcessingTime { get; set; }
        public int TotalQuotes { get; set; }
        public int SuccessfulInjections { get; set; }
        public int FailedInjections { get; set; }
        public int DuplicatesFiltered { get; set; }
        public int DuplicatesDetected { get; set; } = 0;
        public List<MkgQuoteResult> QuoteResults { get; set; } = new List<MkgQuoteResult>();
        public List<string> Errors { get; set; } = new List<string>();
    }

    #endregion

    #region Revision Models

    /// <summary>
    /// MKG Revision Header creation result
    /// </summary>
    public class MkgRevisionHeaderResult
    {
        public bool Success { get; set; }
        public string MkgRevisionNumber { get; set; }
        public string ErrorMessage { get; set; }
        public string HttpStatusCode { get; set; }
        public string RequestPayload { get; set; }
        public string ResponsePayload { get; set; }
    }

    /// <summary>
    /// MKG Revision result
    /// </summary>
    public class MkgRevisionResult
    {
        public string RevisionLineId { get; set; } = "";
        public string ArtiCode { get; set; }
        public string CurrentRevision { get; set; }
        public string NewRevision { get; set; }
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string MkgRevisionId { get; set; }
        public DateTime ProcessedAt { get; set; }
        public string HttpStatusCode { get; set; }
        public List<string> ValidationErrors { get; set; } = new List<string>();
        public string RequestPayload { get; set; }
        public string ResponsePayload { get; set; }

        // Field change tracking
        public string FieldName { get; set; } = "";
        public string FieldChanged { get; set; } = "";
        public string OldValue { get; set; } = "";
        public string NewValue { get; set; } = "";
        public string ChangeReason { get; set; } = "";
    }

    /// <summary>
    /// MKG Revision injection summary
    /// </summary>
    public class MkgRevisionInjectionSummary
    {
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public TimeSpan ProcessingTime { get; set; }
        public int TotalRevisions { get; set; }
        public int SuccessfulInjections { get; set; }
        public int FailedInjections { get; set; }
        public int DuplicatesFiltered { get; set; }
        public int DuplicatesDetected { get; set; } = 0;
        public List<string> Errors { get; set; } = new List<string>();
        public List<MkgRevisionResult> RevisionResults { get; set; } = new List<MkgRevisionResult>();
    }

    /// <summary>
    /// MKG Revision summary (alternate version)
    /// </summary>
    public class MkgRevisionSummary
    {
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public TimeSpan ProcessingTime { get; set; }
        public int TotalRevisions { get; set; }
        public int SuccessfulInjections { get; set; }
        public int FailedInjections { get; set; }
        public int DuplicatesDetected { get; set; } = 0;
        public List<string> Errors { get; set; } = new List<string>();
        public List<MkgRevisionResult> RevisionResults { get; set; } = new List<MkgRevisionResult>();
    }

    #endregion

    #region Testing Models

    /// <summary>
    /// MKG API Test results
    /// </summary>
    public class MkgApiTestResults
    {
        public DateTime TestStartTime { get; set; }
        public DateTime TestEndTime { get; set; }
        public TimeSpan TotalTestTime { get; set; }
        public bool ConfigurationValid { get; set; }
        public bool ApiLoginSuccess { get; set; }
        public bool OrdersWorking { get; set; }
        public int OrdersProcessed { get; set; }
        public bool QuotesWorking { get; set; }
        public int QuotesProcessed { get; set; }
        public bool RevisionsWorking { get; set; }
        public int RevisionsProcessed { get; set; }
        public List<string> TestErrors { get; set; } = new List<string>();
        public List<string> TestLog { get; set; } = new List<string>();
    }

    /// <summary>
    /// Comprehensive MKG test results
    /// </summary>
    public class MkgTestResults
    {
        public bool ConfigurationValid { get; set; }
        public bool ApiLoginSuccess { get; set; }
        public bool OrderControllerHealthy { get; set; }
        public bool QuoteControllerHealthy { get; set; }
        public bool RevisionControllerHealthy { get; set; }
        public bool OrdersWorking { get; set; }
        public bool QuotesWorking { get; set; }
        public bool RevisionsWorking { get; set; }
        public bool CombinedProcessingWorking { get; set; }
        public bool AllSystemsWorking { get; set; }
        public string FullReport { get; set; }
        public TimeSpan TotalTestTime { get; set; }
        public DateTime TestStartTime { get; set; }
        public DateTime TestEndTime { get; set; }
        public int OrdersProcessed { get; set; }
        public int QuotesProcessed { get; set; }
        public int RevisionsProcessed { get; set; }

        public bool AllTestsPassed => ConfigurationValid && ApiLoginSuccess &&
                                    OrderControllerHealthy && QuoteControllerHealthy &&
                                    RevisionControllerHealthy && OrdersWorking &&
                                    QuotesWorking && RevisionsWorking;

        public int WorkingSystemsCount => (OrdersWorking ? 1 : 0) + (QuotesWorking ? 1 : 0) + (RevisionsWorking ? 1 : 0);
    }

    #endregion

    #region Async Injection Support Models

    /// <summary>
    /// Real-time injection display components
    /// </summary>
    public class RealTimeInjectionDisplay
    {
        public DateTime StartTime { get; set; } = DateTime.Now;
        public Label HeaderLabel { get; set; }
        public Panel StatsPanel { get; set; }
        public Label SuccessLabel { get; set; }
        public Label FailedLabel { get; set; }
        public Label TotalLabel { get; set; }
        public Label SpeedLabel { get; set; }
        public Label ElapsedLabel { get; set; }
        public Label ActiveLabel { get; set; }
        public Label QueueLabel { get; set; }
        public Label AvgResponseLabel { get; set; }
        public ProgressBar ProgressBar { get; set; }
        public ListBox ResultsList { get; set; }
        public Button PauseButton { get; set; }
        public Button StopButton { get; set; }
        public Button ExportButton { get; set; }
        public System.Windows.Forms.Timer StatsTimer { get; set; }
    }

    /// <summary>
    /// Injection progress update information
    /// </summary>
    public class InjectionUpdate
    {
        public UpdateType Type { get; set; }
        public string GroupId { get; set; } = "";
        public string ItemId { get; set; } = "";
        public string Status { get; set; } = "";
        public string ErrorMessage { get; set; } = "";
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public long ResponseTime { get; set; }
    }

    /// <summary>
    /// Types of injection updates
    /// </summary>
    public enum UpdateType
    {
        HeaderCreation,
        ItemProcessing,
        ItemSuccess,
        ItemFailure,
        GroupCompleted,
        OverallProgress
    }

    #endregion

    #region Utility Models

    /// <summary>
    /// Saved error result for logging/tracking
    /// </summary>
    public class SavedErrorResult
    {
        public DateTime Timestamp { get; set; }
        public string Content { get; set; }
        public string Type { get; set; }
        public bool IsActualFailure { get; set; }
    }

    #endregion
}