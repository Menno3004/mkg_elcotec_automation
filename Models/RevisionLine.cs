using System;
using System.Globalization;

namespace Mkg_Elcotec_Automation.Models
{
    public class RevisionLine
    {
        private decimal _targetPrice = 0m;

        // Existing properties
        public string LineNumber { get; set; }
        public string ArtiCode { get; set; }
        public string Description { get; set; }
        public string Quantity { get; set; }
        public string Unit { get; set; }
        public string QuotedPrice { get; set; }
        public string CustomerPartNumber { get; set; }
        public string RfqNumber { get; set; }
        public string DrawingNumber { get; set; }
        public string Revision { get; set; }
        public string CurrentRevision { get; set; }
        public string NewRevision { get; set; }
        public string RevisionReason { get; set; }
        public string TechnicalChanges { get; set; }
        public string CommercialChanges { get; set; }
        public string RevisionDate { get; set; }
        public string RevisionStatus { get; set; }
        public string Priority { get; set; }
        public string ApprovalRequired { get; set; }
        public string RequestedDeliveryDate { get; set; }
        public string ExtractionMethod { get; set; }
        public string ExtractionDomain { get; set; }
        public string EmailDomain { get; set; } = "";

        // NEW: Missing properties required for MkgRevisionController
        public string FieldChanged { get; set; } = "";
        public string OldValue { get; set; } = "";
        public string NewValue { get; set; } = "";
        public string ChangeReason { get; set; } = "";

        // Additional properties from code analysis
        public string ClientDomain { get; set; } = "";
        public DateTime ExtractionTimestamp { get; set; } = DateTime.Now;
        public Guid Id { get; set; } = Guid.NewGuid();
        public string OldRevision { get; set; } = "";
        public decimal TotalPrice { get; set; } = 0m;

        public RevisionLine()
        {
            LineNumber = "";
            ArtiCode = "";
            Description = "";
            Quantity = "";
            Unit = "PCS";
            QuotedPrice = "0.00";
            CustomerPartNumber = "";
            RfqNumber = "";
            DrawingNumber = "";
            Revision = "00";
            CurrentRevision = "00";
            NewRevision = "01";
            RevisionReason = "";
            TechnicalChanges = "";
            CommercialChanges = "";
            RevisionDate = "";
            RevisionStatus = "Draft";
            Priority = "Normal";
            ApprovalRequired = "Yes";
            RequestedDeliveryDate = "";
            ExtractionMethod = "";
            ExtractionDomain = "";

            // Initialize new properties
            FieldChanged = "Revision";
            OldValue = CurrentRevision;
            NewValue = NewRevision;
            ChangeReason = RevisionReason;
        }

        public decimal TargetPrice
        {
            get
            {
                if (_targetPrice > 0) return _targetPrice;
                if (decimal.TryParse(QuotedPrice?.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal price))
                    return price;
                return 0m;
            }
            set => _targetPrice = value;
        }

        public void SetExtractionDetails(string method, string domain)
        {
            ExtractionMethod = method;
            ExtractionDomain = domain;
        }

        public void SetEmailDomain(string fromEmail)
        {
            ExtractionDomain = fromEmail ?? "";
            EmailDomain = fromEmail ?? "";
            ClientDomain = fromEmail ?? "";
        }

        /// <summary>
        /// Update the revision change tracking properties when revision values change
        /// </summary>
        public void UpdateRevisionChangeTracking()
        {
            FieldChanged = "Revision";
            OldValue = CurrentRevision;
            NewValue = NewRevision;
            ChangeReason = !string.IsNullOrEmpty(RevisionReason) ? RevisionReason : "Revision update";
        }

        /// <summary>
        /// Set specific field change tracking for detailed revision control
        /// </summary>
        public void SetFieldChange(string fieldName, string oldVal, string newVal, string reason = "")
        {
            FieldChanged = fieldName;
            OldValue = oldVal;
            NewValue = newVal;
            ChangeReason = !string.IsNullOrEmpty(reason) ? reason : $"{fieldName} changed";
        }
    }
}