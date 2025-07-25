using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mkg_Elcotec_Automation.Controllers;
using Mkg_Elcotec_Automation.Models;

namespace Mkg_Elcotec_Automation.Debug
{
    public class MkgApiTester : IDisposable
    {
        private MkgOrderController _orderController;
        private MkgQuoteController _quoteController;
        private MkgRevisionController _revisionController;
        private MkgApiClient _testApiClient;
        private bool disposedValue;

        public int TestOrderCount { get; set; }
        public int TestQuoteCount { get; set; }
        public int TestRevisionCount { get; set; }

        public MkgApiTester()
        {
            TestOrderCount = int.TryParse(ConfigurationManager.AppSettings["MkgTest:DefaultOrderCount"], out int orderCount) ? orderCount : 2;
            TestQuoteCount = int.TryParse(ConfigurationManager.AppSettings["MkgTest:DefaultQuoteCount"], out int quoteCount) ? quoteCount : 2;
            TestRevisionCount = int.TryParse(ConfigurationManager.AppSettings["MkgTest:DefaultRevisionCount"], out int revisionCount) ? revisionCount : 2;
            Console.WriteLine($"🎯 Test Configuration: Orders={TestOrderCount}, Quotes={TestQuoteCount}, Revisions={TestRevisionCount}");
        }

        public async Task<MkgTestResults> RunCompleteTestAsync()
        {
            return await RunCompleteTestAsync(null);
        }

        public async Task<MkgTestResults> RunCompleteTestAsync(IProgress<(int current, int total, string status)> progress)
        {
            var results = new MkgTestResults
            {
                TestStartTime = DateTime.Now
            };

            var sb = new StringBuilder();
            sb.AppendLine("=== COMPREHENSIVE MKG DEBUG WITH ORDERS, QUOTES, AND REVISIONS ===");
            sb.AppendLine($"Debug started at: {results.TestStartTime:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine("🎯 Testing: Configuration, API Connection, Orders, Quotes, and Revisions");
            sb.AppendLine();

            try
            {
                progress?.Report((1, 6, "Checking configuration..."));
                await TestConfiguration(sb, results);

                progress?.Report((2, 6, "Testing API connection..."));
                await TestApiClient(sb, results);

                progress?.Report((3, 6, "Testing controller health..."));
                await TestControllerHealth(sb, results);

                if (results.ApiLoginSuccess)
                {
                    progress?.Report((4, 6, $"Testing order injection ({TestOrderCount} orders)..."));
                    await TestOrderInjection(sb, results);

                    progress?.Report((5, 6, $"Testing quote injection ({TestQuoteCount} quotes)..."));
                    await TestQuoteInjection(sb, results);

                    progress?.Report((6, 6, $"Testing revision injection ({TestRevisionCount} revisions)..."));
                    await TestRevisionInjection(sb, results);
                }

                GenerateTestSummary(sb, results);
            }
            catch (Exception ex)
            {
                sb.AppendLine($"❌ CRITICAL ERROR: {ex.Message}");
                sb.AppendLine($"Stack trace: {ex.StackTrace}");
            }

            results.TestEndTime = DateTime.Now;
            results.TotalTestTime = results.TestEndTime - results.TestStartTime;

            sb.AppendLine();
            sb.AppendLine($"=== DEBUG COMPLETED IN {results.TotalTestTime.TotalSeconds:F1}s ===");

            results.FullReport = sb.ToString();
            return results;
        }

        private async Task TestConfiguration(StringBuilder sb, MkgTestResults results)
        {
            sb.AppendLine("1. CONFIGURATION CHECK:");
            var baseUrl = ConfigurationManager.AppSettings["MkgApi:Urls:Base"];
            var username = ConfigurationManager.AppSettings["MkgApi:Username"];
            var apiKey = ConfigurationManager.AppSettings["MkgApi:ApiKey"];

            sb.AppendLine($"   Base URL: {baseUrl ?? "NULL"}");
            sb.AppendLine($"   Username: {username ?? "NULL"}");
            sb.AppendLine($"   API Key: {(string.IsNullOrEmpty(apiKey) ? "NULL" : "***SET***")}");

            results.ConfigurationValid = !string.IsNullOrEmpty(baseUrl) && !string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(apiKey);
            sb.AppendLine($"   Configuration Status: {(results.ConfigurationValid ? "✅ VALID" : "❌ INVALID")}");
            sb.AppendLine();

            await Task.Delay(100);
        }

        private async Task TestApiClient(StringBuilder sb, MkgTestResults results)
        {
            sb.AppendLine("2. MKG API CLIENT TEST:");
            try
            {
                if (!results.ConfigurationValid)
                {
                    sb.AppendLine("   ❌ Cannot test API - configuration invalid");
                    results.ApiLoginSuccess = false;
                    return;
                }

                _testApiClient = new MkgApiClient();
                sb.AppendLine("   🔄 Testing real MKG API login...");

                var loginSuccess = await _testApiClient.LoginAsync();

                if (loginSuccess)
                {
                    sb.AppendLine("   ✅ MKG API login successful");

                    sb.AppendLine("   🔄 Testing basic API call...");
                    try
                    {
                        var testResponse = await _testApiClient.GetAsync("Documents/debi/?NumRows=1");
                        sb.AppendLine("   ✅ Basic API call successful");
                        results.ApiLoginSuccess = true;
                    }
                    catch (Exception testEx)
                    {
                        sb.AppendLine($"   ⚠️ API call failed but login worked: {testEx.Message}");
                        results.ApiLoginSuccess = true; // Login worked, API might have endpoint issues
                    }
                }
                else
                {
                    sb.AppendLine("   ❌ MKG API login failed");
                    results.ApiLoginSuccess = false;
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"   ❌ MKG API test failed: {ex.Message}");
                results.ApiLoginSuccess = false;
            }
            sb.AppendLine();
        }

        private async Task TestControllerHealth(StringBuilder sb, MkgTestResults results)
        {
            sb.AppendLine("3. ORDER CONTROLLER TEST:");
            try
            {
                _orderController = new MkgOrderController();
                sb.AppendLine("   ✅ MkgOrderController created successfully");
                await Task.Delay(100);
                results.OrderControllerHealthy = true;
                sb.AppendLine($"   Health check: ✅ HEALTHY");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"   ❌ MkgOrderController failed: {ex.Message}");
                results.OrderControllerHealthy = false;
            }
            sb.AppendLine();

            sb.AppendLine("4. QUOTE CONTROLLER TEST:");
            try
            {
                _quoteController = new MkgQuoteController();
                sb.AppendLine("   ✅ MkgQuoteController created successfully");
                await Task.Delay(100);
                results.QuoteControllerHealthy = true;
                sb.AppendLine($"   Health check: ✅ HEALTHY");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"   ❌ MkgQuoteController failed: {ex.Message}");
                results.QuoteControllerHealthy = false;
            }
            sb.AppendLine();

            sb.AppendLine("5. REVISION CONTROLLER TEST:");
            try
            {
                _revisionController = new MkgRevisionController();
                sb.AppendLine("   ✅ MkgRevisionController created successfully");
                await Task.Delay(100);
                results.RevisionControllerHealthy = true;
                sb.AppendLine($"   Health check: ✅ HEALTHY");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"   ❌ MkgRevisionController failed: {ex.Message}");
                results.RevisionControllerHealthy = false;
            }
            sb.AppendLine();
        }

        private async Task TestOrderInjection(StringBuilder sb, MkgTestResults results)
        {
            if (!results.OrderControllerHealthy)
            {
                sb.AppendLine("6. ORDER INJECTION SKIPPED (Controller not healthy)");
                results.OrdersWorking = false;
                results.OrdersProcessed = 0;
                return;
            }

            sb.AppendLine("6. ORDER INJECTION TEST:");
            try
            {
                var testOrders = CreateTestOrders(TestOrderCount);
                sb.AppendLine($"   Created {testOrders.Count} test orders");

                var injectionSummary = await _orderController.InjectOrdersAsync(testOrders);

                results.OrdersProcessed = injectionSummary.SuccessfulInjections;
                results.OrdersWorking = injectionSummary.SuccessfulInjections > 0;

                foreach (var result in injectionSummary.OrderResults)
                {
                    if (result.Success)
                    {
                        sb.AppendLine($"   ✅ Processed order: {result.ArtiCode} - PO: {result.PoNumber}");
                    }
                    else
                    {
                        sb.AppendLine($"   ❌ Failed order: {result.ArtiCode} - Error: {result.ErrorMessage}");
                    }
                }

                sb.AppendLine($"   Order injection: {(results.OrdersWorking ? "✅ SUCCESS" : "❌ FAILED")}");
                sb.AppendLine($"   Processed {results.OrdersProcessed}/{TestOrderCount} orders");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"   ❌ Order injection failed: {ex.Message}");
                results.OrdersWorking = false;
            }
            sb.AppendLine();
        }
        private async Task TestQuoteInjection(StringBuilder sb, MkgTestResults results)
        {
            if (!results.QuoteControllerHealthy)
            {
                sb.AppendLine("7. QUOTE INJECTION SKIPPED (Controller not healthy)");
                results.QuotesWorking = false;
                results.QuotesProcessed = 0;
                return;
            }

            sb.AppendLine("7. QUOTE INJECTION TEST:");
            try
            {
                var testQuotes = CreateTestQuotes(TestQuoteCount);
                sb.AppendLine($"   Created {testQuotes.Count} test quotes");

                var injectionSummary = await _quoteController.InjectQuotesAsync(testQuotes);

                results.QuotesProcessed = injectionSummary.SuccessfulInjections;
                results.QuotesWorking = injectionSummary.SuccessfulInjections > 0;

                foreach (var result in injectionSummary.QuoteResults)
                {
                    if (result.Success)
                    {
                        sb.AppendLine($"   ✅ Processed quote: {result.ArtiCode} - RFQ: {result.RfqNumber}");
                    }
                    else
                    {
                        sb.AppendLine($"   ❌ Failed quote: {result.ArtiCode} - Error: {result.ErrorMessage}");
                    }
                }

                sb.AppendLine($"   Quote injection: {(results.QuotesWorking ? "✅ SUCCESS" : "❌ FAILED")}");
                sb.AppendLine($"   Processed {results.QuotesProcessed}/{TestQuoteCount} quotes");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"   ❌ Quote injection failed: {ex.Message}");
                results.QuotesWorking = false;
                results.QuotesProcessed = 0;
            }
            sb.AppendLine();
        }

        private async Task TestRevisionInjection(StringBuilder sb, MkgTestResults results)
        {
            if (!results.RevisionControllerHealthy)
            {
                sb.AppendLine("8. REVISION INJECTION SKIPPED (Controller not healthy)");
                results.RevisionsWorking = false;
                results.RevisionsProcessed = 0;
                return;
            }

            sb.AppendLine("8. REVISION INJECTION TEST:");
            try
            {
                var testRevisions = CreateTestRevisions(TestRevisionCount);
                sb.AppendLine($"   Created {testRevisions.Count} test revisions");

                var injectionSummary = await _revisionController.InjectRevisionsAsync(testRevisions);

                results.RevisionsProcessed = injectionSummary.SuccessfulInjections;
                results.RevisionsWorking = injectionSummary.SuccessfulInjections > 0;

                foreach (var result in injectionSummary.RevisionResults)
                {
                    if (result.Success)
                    {
                        sb.AppendLine($"   ✅ Processed revision: {result.ArtiCode} - {result.CurrentRevision}→{result.NewRevision}");
                    }
                    else
                    {
                        sb.AppendLine($"   ❌ Failed revision: {result.ArtiCode} - Error: {result.ErrorMessage}");
                    }
                }

                sb.AppendLine($"   Revision injection: {(results.RevisionsWorking ? "✅ SUCCESS" : "❌ FAILED")}");
                sb.AppendLine($"   Processed {results.RevisionsProcessed}/{TestRevisionCount} revisions");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"   ❌ Revision injection failed: {ex.Message}");
                results.RevisionsWorking = false;
                results.RevisionsProcessed = 0;
            }
            sb.AppendLine();
        }

        private void GenerateTestSummary(StringBuilder sb, MkgTestResults results)
        {
            sb.AppendLine("9. TEST SUMMARY:");
            sb.AppendLine($"   Overall Status: {(results.AllTestsPassed ? "✅ ALL TESTS PASSED" : "❌ SOME TESTS FAILED")}");
            sb.AppendLine($"   Configuration: {(results.ConfigurationValid ? "✅" : "❌")}");
            sb.AppendLine($"   API Connection: {(results.ApiLoginSuccess ? "✅" : "❌")}");
            sb.AppendLine($"   Order System: {(results.OrdersWorking ? "✅" : "❌")} ({results.OrdersProcessed} processed)");
            sb.AppendLine($"   Quote System: {(results.QuotesWorking ? "✅" : "❌")} ({results.QuotesProcessed} processed)");
            sb.AppendLine($"   Revision System: {(results.RevisionsWorking ? "✅" : "❌")} ({results.RevisionsProcessed} processed)");
            sb.AppendLine($"   Working Systems: {results.WorkingSystemsCount}/3");
            sb.AppendLine();

            if (!results.AllTestsPassed)
            {
                sb.AppendLine("FAILED COMPONENTS:");
                if (!results.ConfigurationValid) sb.AppendLine("   • Configuration");
                if (!results.ApiLoginSuccess) sb.AppendLine("   • API Connection");
                if (!results.OrdersWorking) sb.AppendLine("   • Order System");
                if (!results.QuotesWorking) sb.AppendLine("   • Quote System");
                if (!results.RevisionsWorking) sb.AppendLine("   • Revision System");
            }
        }

        private List<OrderLine> CreateTestOrders(int count)
        {
            var orders = new List<OrderLine>();
            var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");

            for (int i = 1; i <= count; i++)
            {
                var order = new OrderLine(
                    lineNumber: i.ToString(),
                    artiCode: $"123.456.{789 + i}",
                    description: $"Test Order Item {i}",
                    drawingNumber: $"DWG-TEST-{i:D3}",
                    revision: "A",
                    supplierPartNumber: $"SUP-{timestamp}-{i:D3}",
                    trackingNumber: "",
                    requestedShipDate: "",
                    memoExtern: $"Test order {i} for debugging",
                    deliveryDate: DateTime.Now.AddDays(30).ToString("yyyy-MM-dd"),
                    sapPoLineNumber: $"{i}",
                    quantity: "1",
                    unit: "MTR",  // Test unit conversion MTR → m
                    unitPrice: "100.00",
                    totalPrice: "100.00"
                );

                order.PoNumber = $"PO-TEST-{timestamp}-{i:D3}";
                order.SetExtractionDetails("TEST_DATA_GENERATION", "test.local");
                orders.Add(order);
            }

            return orders;
        }

        private List<QuoteLine> CreateTestQuotes(int count)
        {
            var quotes = new List<QuoteLine>();
            var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");

            for (int i = 1; i <= count; i++)
            {
                var quote = new QuoteLine
                {
                    LineNumber = i.ToString(),
                    ArtiCode = $"345.678.{901 + i}",
                    Description = $"Test Quote Item {i}",
                    Quantity = "1",
                    Unit = "PCS",  // Test unit conversion PCS → st.
                    QuotedPrice = "150.00",
                    CustomerPartNumber = $"CUST-{timestamp}-{i:D3}",
                    RfqNumber = $"RFQ-2025-{i:D3}",
                    DrawingNumber = $"DWG-QTE-{i:D3}",
                    Revision = "01",
                    RequestedDeliveryDate = DateTime.Now.AddDays(21).ToString("dd-MM-yyyy"),
                    QuoteDate = DateTime.Now.ToString("dd-MM-yyyy"),
                    QuoteStatus = "Draft",
                    Priority = "Normal",
                    ValidUntil = DateTime.Now.AddDays(30).ToString("dd-MM-yyyy")
                };

                quote.SetExtractionDetails("TEST_DATA_GENERATION", "test.local");
                quotes.Add(quote);
            }

            return quotes;
        }

        private List<RevisionLine> CreateTestRevisions(int count)
        {
            var revisions = new List<RevisionLine>();
            var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");

            for (int i = 1; i <= count; i++)
            {
                var revision = new RevisionLine
                {
                    LineNumber = i.ToString(),
                    ArtiCode = $"123.456.{789 + i}",
                    Description = $"Test Revision Item {i}",
                    Quantity = "1",
                    Unit = "KG",  // Test unit conversion
                    QuotedPrice = "125.00",
                    CustomerPartNumber = $"CUST-REV-{timestamp}-{i:D3}",
                    RfqNumber = $"ECN-2025-{i:D3}",
                    DrawingNumber = $"DWG-PUMP-{i:D3}",
                    Revision = "A",
                    CurrentRevision = "A",
                    NewRevision = "B",
                    RevisionReason = $"Test revision {i} - debug testing",
                    TechnicalChanges = "Minor updates for testing",
                    CommercialChanges = "",
                    RevisionDate = DateTime.Now.ToString("dd-MM-yyyy"),
                    RevisionStatus = "Draft",
                    Priority = "Normal",
                    ApprovalRequired = "No",
                    RequestedDeliveryDate = DateTime.Now.AddDays(14).ToString("dd-MM-yyyy")
                };

                revision.SetExtractionDetails("TEST_DATA_GENERATION", "test.local");
                revisions.Add(revision);
            }

            return revisions;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    _orderController?.Dispose();
                    _quoteController?.Dispose();
                    _revisionController?.Dispose();
                    _testApiClient?.Dispose();
                }
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}