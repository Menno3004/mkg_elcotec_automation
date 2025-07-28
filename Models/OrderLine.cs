using System;

namespace Mkg_Elcotec_Automation.Models
{
    public class OrderLine
    {
        public string LineNumber { get; set; }
        public string ArtiCode { get; set; }
        public string Description { get; set; }
        public string DrawingNumber { get; set; }
        public string Revision { get; set; }
        public string SupplierPartNumber { get; set; }
        public string TrackingNumber { get; set; }
        public string RequestedShipDate { get; set; }
        public string MemoExtern { get; set; }
        public string DeliveryDate { get; set; }
        public string SapPoLineNumber { get; set; }
        public string Quantity { get; set; }
        public string Unit { get; set; }
        public string UnitPrice { get; set; }
        public string TotalPrice { get; set; }
        public string PoNumber { get; set; }
        public string PoDate { get; set; }
        public string Notes { get; set; }
        public string ExtractionMethod { get; set; } = "UNKNOWN";
        public string ExtractionSource { get; set; } = "";
        public string ExtractionDomain { get; set; } = "";
        public DateTime ExtractedAt { get; set; } = DateTime.Now;
        public string EmailDomain { get; set; } = "";
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string DebtorNumber { get; set; }
        public string RequestedDeliveryDate { get; internal set; }
        public string OrderStatus { get; internal set; }
        public string CustomerPartNumber { get; internal set; }
        public string Priority { get; internal set; }
        public string DeliveryTerms { get; internal set; }
        public string OrderDate { get; internal set; }
        public string InternalReference { get; internal set; }
        public string ApprovalRequired { get; internal set; }
        public string ExpressHandling { get; internal set; }
        public string SpecialHandling { get; internal set; }
        public string ShippingMethod { get; internal set; }
        public string DeliveryAddress { get; internal set; }
        public string DocumentationRequired { get; internal set; }
        public string ConfirmedDeliveryDate { get; internal set; }
        public string QualityRequirements { get; internal set; }
        public string ShippedDate { get; internal set; }
        public string DiscountApplied { get; internal set; }

        public string ValidationStatus { get; set; } = "PENDING";
        public string ValidationNotes { get; set; } = "";
        public decimal? MkgReferencePrice { get; set; }
        public DateTime? ValidationTimestamp { get; set; }

        public string ValidatedBy { get; set; } = "SYSTEM";
        public OrderLine()
        {
            LineNumber = "";
            ArtiCode = "";
            Description = "";
            DrawingNumber = "";
            Revision = "00";
            SupplierPartNumber = "";
            TrackingNumber = "";
            RequestedShipDate = "";
            MemoExtern = "";
            DeliveryDate = "";
            SapPoLineNumber = "";
            Quantity = "";
            Unit = "PCS";
            UnitPrice = "0.00";
            TotalPrice = "0.00";
            PoNumber = "";
            PoDate = "";
            Notes = "";
            DebtorNumber = "";
        }

        public OrderLine(string lineNumber, string artiCode, string description, string drawingNumber,
                        string revision, string supplierPartNumber, string trackingNumber, string requestedShipDate,
                        string memoExtern, string deliveryDate, string sapPoLineNumber, string quantity,
                        string unit, string unitPrice, string totalPrice)
        {
            LineNumber = lineNumber ?? "";
            ArtiCode = artiCode ?? "";
            Description = description ?? "";
            DrawingNumber = drawingNumber ?? "";
            Revision = revision ?? "00";
            SupplierPartNumber = supplierPartNumber ?? "";
            TrackingNumber = trackingNumber ?? "";
            RequestedShipDate = requestedShipDate ?? "";
            MemoExtern = memoExtern ?? "";
            DeliveryDate = deliveryDate ?? "";
            SapPoLineNumber = sapPoLineNumber ?? "";
            Quantity = quantity ?? "";
            Unit = unit ?? "PCS";
            UnitPrice = unitPrice ?? "0.00";
            TotalPrice = totalPrice ?? "0.00";
            PoNumber = "";
            PoDate = "";
            Notes = "";
            ExtractionMethod = "UNKNOWN";
            ExtractionSource = "";
            ExtractionDomain = "";
            ExtractedAt = DateTime.Now;
            EmailDomain = "";
            Id = Guid.NewGuid().ToString();
            DebtorNumber = "";
        }

        public void SetExtractionDetails(string method, string domain)
        {
            ExtractionMethod = method;
            ExtractionDomain = domain;
            ExtractedAt = DateTime.Now;
        }
    }
}