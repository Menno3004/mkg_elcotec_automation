using System;
using System.Globalization;

namespace Mkg_Elcotec_Automation.Models
{
    public class QuoteLine
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
        public string RequestedDeliveryDate { get; set; }
        public string QuoteDate { get; set; }
        public string QuoteStatus { get; set; }
        public string Priority { get; set; }
        public string ValidUntil { get; set; }
        public string ExtractionMethod { get; set; }
        public string ExtractionDomain { get; set; }
        public string EmailDomain { get; set; } = "";

        // NEW: Missing properties required for MkgQuoteController
        public string QuoteNotes { get; set; } = "";
        public string QuoteValidUntil { get; set; } = "";

        // Additional properties for completeness
        public string ClientDomain { get; set; } = "";
        public DateTime ExtractionTimestamp { get; set; } = DateTime.Now;
        public Guid Id { get; set; } = Guid.NewGuid();
        public decimal TotalPrice { get; set; } = 0m;
        public string SpecialInstructions { get; set; } = "";
        public string PaymentTerms { get; set; } = "";
        public string DeliveryTerms { get; set; } = "";

        // Default constructor
        public QuoteLine()
        {
            LineNumber = "";
            ArtiCode = "";
            Description = "";
            Quantity = "1";
            Unit = "PCS";
            QuotedPrice = "0.00";
            CustomerPartNumber = "";
            RfqNumber = "";
            DrawingNumber = "";
            Revision = "00";
            RequestedDeliveryDate = "";
            QuoteDate = DateTime.Now.ToString("dd-MM-yyyy");
            QuoteStatus = "Draft";
            Priority = "Normal";
            ValidUntil = DateTime.Now.AddDays(30).ToString("dd-MM-yyyy");
            ExtractionMethod = "";
            ExtractionDomain = "";

            // Initialize new properties
            QuoteNotes = "";
            QuoteValidUntil = ValidUntil; // Mirror ValidUntil for backwards compatibility
        }

        // ✅ FIXED: Constructor with DEFAULT PARAMETERS to support named parameter syntax
        public QuoteLine(
            string lineNumber = "",
            string artiCode = "",
            string description = "",
            string quantity = "1",
            string unit = "PCS",
            string quotedPrice = "0.00",
            string customerPartNumber = "",
            string rfqNumber = "",
            string drawingNumber = "",
            string revision = "00",
            string requestedDeliveryDate = "",
            string quoteDate = "",
            string quoteStatus = "Draft",
            string priority = "Normal",
            string validUntil = "",
            string extractionMethod = "",
            string extractionDomain = "")
        {
            LineNumber = lineNumber ?? "";
            ArtiCode = artiCode ?? "";
            Description = description ?? "";
            Quantity = quantity ?? "1";
            Unit = unit ?? "PCS";
            QuotedPrice = quotedPrice ?? "0.00";
            CustomerPartNumber = customerPartNumber ?? "";
            RfqNumber = rfqNumber ?? "";
            DrawingNumber = drawingNumber ?? "";
            Revision = revision ?? "00";
            RequestedDeliveryDate = requestedDeliveryDate ?? "";
            QuoteDate = !string.IsNullOrEmpty(quoteDate) ? quoteDate : DateTime.Now.ToString("dd-MM-yyyy");
            QuoteStatus = quoteStatus ?? "Draft";
            Priority = priority ?? "Normal";
            ValidUntil = !string.IsNullOrEmpty(validUntil) ? validUntil : DateTime.Now.AddDays(30).ToString("dd-MM-yyyy");
            ExtractionMethod = extractionMethod ?? "";
            ExtractionDomain = extractionDomain ?? "";

            // Initialize additional properties
            EmailDomain = extractionDomain ?? "";
            ClientDomain = extractionDomain ?? "";
            QuoteValidUntil = ValidUntil;
            QuoteNotes = "";
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

        public dynamic Currency { get; internal set; }
        public dynamic QuoteReference { get; internal set; }
        public dynamic LeadTime { get; internal set; }
        public string Notes { get; internal set; }

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
        /// Calculate total price based on quantity and quoted price
        /// </summary>
        public void CalculateTotalPrice()
        {
            if (decimal.TryParse(Quantity, out decimal qty) &&
                decimal.TryParse(QuotedPrice?.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal price))
            {
                TotalPrice = qty * price;
            }
        }

        /// <summary>
        /// Set quote validity period in days from now
        /// </summary>
        public void SetValidityPeriod(int days)
        {
            var validDate = DateTime.Now.AddDays(days);
            ValidUntil = validDate.ToString("dd-MM-yyyy");
            QuoteValidUntil = validDate.ToString("yyyy-MM-dd"); // MKG format
        }

        /// <summary>
        /// Update quote notes and special instructions
        /// </summary>
        public void SetQuoteDetails(string notes, string specialInstructions = "")
        {
            QuoteNotes = notes ?? "";
            SpecialInstructions = specialInstructions ?? "";
        }
    }
}