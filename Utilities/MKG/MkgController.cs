using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Text.Json;
using Mkg_Elcotec_Automation.Models;
using System.Linq;

namespace Mkg_Elcotec_Automation.Controllers
{
    /// <summary>
    /// Base controller - minimal shared functionality only
    /// </summary>
    public class MkgController : IDisposable
    {
        protected readonly MkgApiClient _mkgApiClient;
        private static readonly Dictionary<string, CustomerInfo> _customerCache = new Dictionary<string, CustomerInfo>();
        private static readonly object _cacheLock = new object();
        private const int CACHE_DURATION_HOURS = 24;

        public MkgController()
        {
            _mkgApiClient = new MkgApiClient();
        }
        #region Utility Methods

        /// <summary>
        /// Convert common unit variations to valid MKG units
        /// This ensures API compatibility and prevents unit-related errors
        /// </summary>
        protected string ConvertToValidUnit(string unit)
        {
            if (string.IsNullOrEmpty(unit))
                return "st."; // Default to pieces if no unit specified
                return "st."; // Default to pieces if no unit specified

            // Convert common invalid units to valid MKG units
            switch (unit.ToUpper().Trim())
            {
                // Pieces/Items
                case "ST":
                case "STK":
                case "STUK":
                case "STUKS":
                case "PIECE":
                case "PIECES":
                case "PC":
                case "PCS":
                case "EA":
                case "EACH":
                case "ITEM":
                case "ITEMS":
                    return "st."; //this opne

                // Length units
                case "M":
                case "MTR":
                case "METER":
                case "METRES":
                case "METERS":
                case "METRE":
                    return "m";

                case "MM":
                case "MILLIMETER":
                case "MILLIMETRE":
                    return "mm";

                case "CM":
                case "CENTIMETER":
                case "CENTIMETRE":
                    return "cm";

                // Weight units
                case "KG":
                case "KILOGRAM":
                case "KILOGRAMS":
                    return "kg";

                case "G":
                case "GRAM":
                case "GRAMS":
                    return "g";

                case "T":
                case "TON":
                case "TONS":
                case "TONNE":
                case "TONNES":
                    return "t";

                // Volume units
                case "L":
                case "LITER":
                case "LITRE":
                case "LITTERS":
                case "LITRES":
                    return "l";

                case "ML":
                case "MILLILITER":
                case "MILLILITRE":
                    return "ml";

                // Area units
                case "M2":
                case "M²":
                case "SQM":
                case "SQUARE METER":
                case "SQUARE METRE":
                    return "m²";

                // Time units
                case "H":
                case "HR":
                case "HOUR":
                case "HOURS":
                case "UUR":
                case "UREN":
                    return "uur";

                // Special cases for common variations
                case "SET":
                case "SETS":
                    return "set";

                case "PAIR":
                case "PAIRS":
                case "PR":
                    return "paar";

                case "PACK":
                case "PACKAGE":
                case "PACKAGES":
                case "PKG":
                    return "pak";

                case "BOX":
                case "BOXES":
                    return "doos";

                default:
                    // Log unusual units for review
                    LogDebug($"⚠️ Unknown unit '{unit}' - using as-is (lowercase)");
                    return unit.ToLower();
            }
        }

        /// <summary>
        /// Validate that a unit is acceptable for MKG
        /// Returns true if unit is valid, false if it might cause issues
        /// </summary>
        protected bool IsValidMkgUnit(string unit)
        {
            if (string.IsNullOrEmpty(unit))
                return false;

            var validUnits = new HashSet<string>
            {
                "st.", "m", "mm", "cm", "kg", "g", "t", "l", "ml", "m²",
                "uur", "set", "paar", "pak", "doos"
            };

            return validUnits.Contains(unit.ToLower());
        }

        /// <summary>
        /// Get a suggested valid unit for an invalid unit
        /// </summary>
        protected string GetSuggestedUnit(string invalidUnit)
        {
            if (string.IsNullOrEmpty(invalidUnit))
                return "st.";

            var converted = ConvertToValidUnit(invalidUnit);
            return IsValidMkgUnit(converted) ? converted : "st.";
        }
        protected List<OrderLine> FilterValidOrderLines(List<OrderLine> orders)
        {
            LogDebug($"🔄 Filtering {orders.Count} order lines for validity");

            var validOrders = orders.Where(order =>
                !string.IsNullOrWhiteSpace(order.ArtiCode) &&
                !string.IsNullOrWhiteSpace(order.PoNumber) &&
                !order.ArtiCode.StartsWith("QUOTE-ITEM-", StringComparison.OrdinalIgnoreCase) &&
                !order.ArtiCode.StartsWith("RFQ-", StringComparison.OrdinalIgnoreCase) &&
                !order.ArtiCode.StartsWith("UNKNOWN-ARTICLE-", StringComparison.OrdinalIgnoreCase) &&
                !order.ArtiCode.Equals("NUMBER", StringComparison.OrdinalIgnoreCase) &&
                !order.ArtiCode.EndsWith("DELIVER", StringComparison.OrdinalIgnoreCase) &&
                !IsInvalidPoNumber(order.PoNumber)
            ).ToList();

            LogDebug($"✅ Filtered to {validOrders.Count} valid order lines");
            return validOrders;
        }

        private bool IsInvalidPoNumber(string poNumber)
        {
            return poNumber == "PO-0" ||
                   poNumber == "INVALID" ||
                   (poNumber != null && poNumber.Contains("--"));
        }

        protected async Task<CustomerInfo> FindCustomerByEmailDomain(string emailAddress)
        {
            if (string.IsNullOrEmpty(emailAddress))
                return GetDefaultCustomer();

            try
            {
                var domain = ExtractDomainFromEmail(emailAddress);
                if (string.IsNullOrEmpty(domain))
                    return GetDefaultCustomer();

                // Check cache first
                lock (_cacheLock)
                {
                    if (_customerCache.TryGetValue(domain, out var cachedCustomer))
                    {
                        if (DateTime.Now.Subtract(cachedCustomer.CachedAt).TotalHours < CACHE_DURATION_HOURS)
                        {
                            return cachedCustomer;
                        }
                        _customerCache.Remove(domain);
                    }
                }

                // Search in MKG
                var customer = await SearchCustomerInMkg(domain);
                if (customer != null)
                {
                    // Cache the result
                    lock (_cacheLock)
                    {
                        customer.CachedAt = DateTime.Now;
                        _customerCache[domain] = customer;
                    }
                    return customer;
                }

                return GetDefaultCustomer();
            }
            catch
            {
                return GetDefaultCustomer();
            }
        }

        private string ExtractDomainFromEmail(string emailAddress)
        {
            try
            {
                var atIndex = emailAddress.IndexOf('@');
                if (atIndex > 0 && atIndex < emailAddress.Length - 1)
                {
                    return emailAddress.Substring(atIndex + 1).ToLower();
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        private async Task<CustomerInfo> SearchCustomerInMkg(string domain)
        {
            try
            {
                var endpoint = $"Documents/debi/?Filter=debi_actief = true AND debi_naam CONTAINS \"{domain}\"&FieldList=admi_num,debi_num,rela_num,debi_naam,debi_actief&NumRows=1";
                var response = await _mkgApiClient.GetAsync(endpoint);

                if (string.IsNullOrEmpty(response))
                    return null;

                var jsonDoc = JsonDocument.Parse(response);
                var root = jsonDoc.RootElement;

                if (root.TryGetProperty("data", out var dataElement) && dataElement.GetArrayLength() > 0)
                {
                    var customer = dataElement[0];
                    return new CustomerInfo
                    {
                        CustomerName = GetStringValue(customer, "debi_naam"),
                        AdministrationNumber = GetIntValue(customer, "admi_num").ToString(),
                        DebtorNumber = GetIntValue(customer, "debi_num").ToString(),
                        RelationNumber = GetIntValue(customer, "rela_num").ToString(),
                        IsActive = GetBoolValue(customer, "debi_actief"),
                        EmailDomain = domain
                    };
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        private CustomerInfo GetDefaultCustomer()
        {
            return new CustomerInfo
            {
                CustomerName = "Default Customer",
                AdministrationNumber = "1",
                DebtorNumber = "30010",
                RelationNumber = "2",
                IsActive = true,
                EmailDomain = ""
            };
        }

        private int GetIntValue(JsonElement element, string propertyName)
        {
            if (element.TryGetProperty(propertyName, out var prop))
            {
                if (prop.ValueKind == JsonValueKind.Number)
                    return prop.GetInt32();
                if (prop.ValueKind == JsonValueKind.String && int.TryParse(prop.GetString(), out var intValue))
                    return intValue;
            }
            return 0;
        }

        private bool GetBoolValue(JsonElement element, string propertyName)
        {
            if (element.TryGetProperty(propertyName, out var prop))
            {
                if (prop.ValueKind == JsonValueKind.True) return true;
                if (prop.ValueKind == JsonValueKind.False) return false;
            }
            return false;
        }

        private string GetStringValue(JsonElement element, string propertyName)
        {
            if (element.TryGetProperty(propertyName, out var prop))
            {
                return prop.GetString() ?? "";
            }
            return "";
        }

        protected void LogDebug(string message)
        {
            Console.WriteLine($"[MkgController] {DateTime.Now:HH:mm:ss} {message}");
        }

        public virtual void Dispose()
        {
            _mkgApiClient?.Dispose();
        }
        #endregion
    }
}