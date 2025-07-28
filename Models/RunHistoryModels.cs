using System;
using System.Collections.Generic;

namespace Mkg_Elcotec_Automation.Models
{
    /// <summary>
    /// Represents a single automation run for history tracking
    /// </summary>
    public class RunHistoryItem
    {
        public Guid RunId { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime? EndTime { get; set; }
        public string Status { get; set; } = "Running";
        public int EmailsProcessed { get; set; }
        public int OrdersFound { get; set; }
        public int QuotesFound { get; set; }
        public int RevisionsFound { get; set; }
        public int ErrorsEncountered { get; set; }
        public int TotalItems { get; set; }
        public bool IsCurrentRun { get; set; }
        public Dictionary<string, object> Settings { get; set; } = new();

        // 🎯 NEW: Add injection failure tracking properties
        public bool HasInjectionFailures { get; set; } = false;
        public int TotalFailuresAtCompletion { get; set; } = 0;

        public decimal SuccessRate => TotalItems > 0 ?
            ((decimal)(TotalItems - ErrorsEncountered) / TotalItems) * 100 : 0;

        public TimeSpan Duration => EndTime.HasValue ?
            EndTime.Value - StartTime : DateTime.Now - StartTime;

        public string DisplayText => IsCurrentRun
            ? "⏳ Current Run (Live)"
             : $"{StatusIcon} {StartTime:MM-dd HH:mm} | {SuccessRate:F0}% | O:{OrdersFound} Q:{QuotesFound} R:{RevisionsFound}";
        // 🔧 FIXED: Improved StatusIcon logic using HasInjectionFailures
        private string StatusIcon
        {
            get
            {
                // For current/running status
                if (Status == "Running" || IsCurrentRun)
                    return "⏳";

                if (Status == "Failed")
                    return "❌";

                if (Status == "Cancelled")
                    return "🚫";

                // For completed runs, check injection failures first, then success rate
                if (Status == "Completed")
                {
                    // If we have injection failure data, use that (most accurate)
                    if (TotalFailuresAtCompletion > 0 || HasInjectionFailures)
                        return "🔴"; // Red for any injection failures

                    // If no injection failures, check overall success rate
                    if (SuccessRate >= 100)
                        return "🟢"; // Green for perfect success
                    else if (SuccessRate >= 90)
                        return "🟡"; // Yellow for mostly successful
                    else
                        return "🔴"; // Red for low success rate
                }

                // Default fallback
                return "⚪"; // Grey for unknown status
            }
        }

    }
}