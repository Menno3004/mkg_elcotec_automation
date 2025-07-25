using System;
using System.Globalization;
using System.IO;
using HtmlAgilityPack;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Linq;
using Mkg_Elcotec_Automation.Models;

namespace Mkg_Elcotec_Automation
{
    class OrderHtmlParser
    {
        public static HtmlDocument Doc { get; set; }

        private static List<string> DebugLog = new List<string>();

        public static void ClearDebugLog() => DebugLog.Clear();
        public static List<string> GetDebugLog() => new List<string>(DebugLog);

        private static void LogDebug(string message)
        {
            var logEntry = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
            DebugLog.Add(logEntry);
            Console.WriteLine($"[HTML_PARSER] {logEntry}");
        }
        public static HtmlDocument LoadHtml(string body, string folder)
        {
            LogDebug($"Loading HTML content, length: {body?.Length ?? 0}");

            if (string.IsNullOrWhiteSpace(body))
            {
                LogDebug("ERROR: HTML body is null or empty");
                return null;
            }

            try
            {
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(body);
                Doc = doc;

                // Debug: Check document structure
                LogDebug($"HTML document loaded successfully");
                LogDebug($"Document has {doc.DocumentNode?.ChildNodes?.Count ?? 0} child nodes");

                // Check for order_lines table
                var orderTable = doc.DocumentNode?.SelectSingleNode("//table[@id='order_lines']");
                LogDebug($"Order lines table found: {orderTable != null}");

                if (orderTable != null)
                {
                    var rows = orderTable.SelectNodes(".//tr");
                    LogDebug($"Order table has {rows?.Count ?? 0} rows");
                }

                // Check for other potential order tables
                var allTables = doc.DocumentNode?.SelectNodes("//table");
                LogDebug($"Total tables in document: {allTables?.Count ?? 0}");

                if (allTables != null)
                {
                    for (int i = 0; i < allTables.Count; i++)
                    {
                        var table = allTables[i];
                        var tableId = table.GetAttributeValue("id", "no-id");
                        var tableClass = table.GetAttributeValue("class", "no-class");
                        LogDebug($"Table {i + 1}: id='{tableId}', class='{tableClass}'");
                    }
                }

                return doc;
            }
            catch (Exception ex)
            {
                LogDebug($"ERROR loading HTML: {ex.Message}");
                return null;
            }
        }
        private Dictionary<string, int> AnalyzeTableStructure(HtmlNode table)
        {
            var columnMapping = new Dictionary<string, int>
    {
        {"UnitPrice", 6},    // Default positions from your original code
        {"TotalPrice", 7}
    };

            try
            {
                var rows = table.SelectNodes(".//tr");
                if (rows == null) return columnMapping;

                // Look for header row
                foreach (var row in rows)
                {
                    var headers = row.SelectNodes(".//th");
                    if (headers != null && headers.Count > 0)
                    {
                        // Found header row, analyze column names
                        for (int i = 0; i < headers.Count; i++)
                        {
                            var headerText = headers[i].InnerText?.ToLower().Trim() ?? "";
                            Console.WriteLine($"📋 Header {i}: '{headerText}'");

                            if (headerText.Contains("unit") && headerText.Contains("price"))
                                columnMapping["UnitPrice"] = i;
                            else if (headerText.Contains("price") && !headerText.Contains("total") && !columnMapping.ContainsKey("UnitPrice"))
                                columnMapping["UnitPrice"] = i;
                            else if (headerText.Contains("total") || headerText.Contains("extended") || headerText.Contains("amount"))
                                columnMapping["TotalPrice"] = i;
                        }
                        break; // Found headers, stop looking
                    }
                }

                // If no headers found, try to detect from data patterns
                if (!columnMapping.ContainsKey("UnitPrice") || columnMapping["UnitPrice"] == 6)
                {
                    DetectPriceColumnsFromData(table, columnMapping);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Error analyzing table structure: {ex.Message}");
            }

            return columnMapping;
        }

        /// <summary>
        /// Detect price columns by analyzing data patterns when headers are unclear
        /// </summary>
        private void DetectPriceColumnsFromData(HtmlNode table, Dictionary<string, int> columnMapping)
        {
            try
            {
                var rows = table.SelectNodes(".//tr");
                if (rows == null) return;

                var pricePatterns = new Regex(@"€\s*\d+[.,]\d{2}|\$\s*\d+[.,]\d{2}|\d+[.,]\d{2}\s*€|\d+[.,]\d{2}\s*\$");
                var columnPriceCount = new Dictionary<int, int>();

                // Analyze first 5 data rows to find which columns contain prices
                int rowsAnalyzed = 0;
                foreach (var row in rows)
                {
                    var columns = row.SelectNodes(".//td");
                    if (columns == null || rowsAnalyzed >= 5) continue;

                    for (int i = 0; i < columns.Count; i++)
                    {
                        var cellText = columns[i].InnerText?.Trim() ?? "";
                        if (pricePatterns.IsMatch(cellText))
                        {
                            columnPriceCount[i] = columnPriceCount.GetValueOrDefault(i, 0) + 1;
                        }
                    }
                    rowsAnalyzed++;
                }

                // Update mapping based on detected patterns
                var priceColumns = columnPriceCount.Where(kv => kv.Value >= 2).OrderBy(kv => kv.Key).ToList();
                if (priceColumns.Count >= 1)
                {
                    columnMapping["UnitPrice"] = priceColumns[0].Key;
                    Console.WriteLine($"🔍 Detected UnitPrice column: {priceColumns[0].Key}");
                }
                if (priceColumns.Count >= 2)
                {
                    columnMapping["TotalPrice"] = priceColumns[1].Key;
                    Console.WriteLine($"🔍 Detected TotalPrice column: {priceColumns[1].Key}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Error detecting price columns: {ex.Message}");
            }
        }

        /// <summary>
        /// Extract basic order data using your existing logic
        /// </summary>
        private (string artiCode, string description, string drawingNumber, string revision,
                 string supplierPartNumber, string requestedShippingDate, string articleDescription,
                 string sapPoLineNumber) ExtractBasicOrderData(string dataString, HtmlNodeCollection columns)
        {
            // Your existing extraction logic (unchanged for compatibility)
            int descriptionIndex = dataString.IndexOf("<");
            string description = dataString.Substring(0, descriptionIndex);
            string artiCode = description.Substring(0, description.IndexOf(" "));
            if (artiCode.Contains('A'))
            {
                artiCode = artiCode.Substring(0, artiCode.IndexOf('A'));
            }

            // Extract drawing number and revision (your existing logic)
            int drawingNumberIndex = dataString.IndexOf("Drawing Number:");
            if (drawingNumberIndex == -1)
            {
                drawingNumberIndex = dataString.IndexOf("Drawing No:");
                if (drawingNumberIndex == -1)
                {
                    return (artiCode, description, "", "", "", "", "", "");
                }
            }

            dataString = dataString.Substring(drawingNumberIndex);
            string drawingNumberRaw = dataString.Substring(dataString.IndexOf(":"), dataString.IndexOf("<") + 1 - dataString.IndexOf(":"));
            int revIndex = drawingNumberRaw.IndexOf("rev.");
            if (revIndex == -1)
            {
                revIndex = drawingNumberRaw.IndexOf("Rev.");
            }

            string drawingNumber = drawingNumberRaw.Substring(0, revIndex);
            drawingNumber = drawingNumber.Substring(1);
            string revisionRaw = drawingNumberRaw.Substring(revIndex);
            string revision = revisionRaw.Substring(revisionRaw.IndexOf(" "), 2).Trim();

            // Extract other fields (simplified version of your logic)
            string supplierPartNumber = ExtractFieldValue(dataString, "Supplier Part Number:");
            string requestedShippingDate = ExtractFieldValue(dataString, "Requested Ship Date:");
            string sapPoLineNumber = ExtractFieldValue(dataString, "SAP PO Line Number:");

            // Article description extraction (simplified)
            string articleDescription = ""; // Implement your existing complex logic here if needed

            return (artiCode, description, drawingNumber, revision, supplierPartNumber,
                    requestedShippingDate, articleDescription, sapPoLineNumber);
        }

        /// <summary>
        /// Enhanced price extraction with dynamic column mapping
        /// </summary>
        private (string unitPrice, string totalPrice) ExtractPricesWithDynamicMapping(
            HtmlNodeCollection columns, Dictionary<string, int> columnMapping, string artiCode)
        {
            string unitPrice = "0.00";
            string totalPrice = "0.00";

            try
            {
                // Method 1: Use dynamic column mapping
                int unitPriceColumn = columnMapping.GetValueOrDefault("UnitPrice", 6);
                int totalPriceColumn = columnMapping.GetValueOrDefault("TotalPrice", 7);

                if (unitPriceColumn < columns.Count)
                {
                    var unitPriceText = columns[unitPriceColumn].InnerText?.Trim() ?? "";
                    unitPrice = ExtractPriceFromString(unitPriceText);
                    Console.WriteLine($"💰 Unit price from column {unitPriceColumn}: '{unitPriceText}' -> {unitPrice}");
                }

                if (totalPriceColumn < columns.Count)
                {
                    var totalPriceText = columns[totalPriceColumn].InnerText?.Trim() ?? "";
                    totalPrice = ExtractPriceFromString(totalPriceText);
                    Console.WriteLine($"💰 Total price from column {totalPriceColumn}: '{totalPriceText}' -> {totalPrice}");
                }

                // Method 2: Fallback - scan all columns for price patterns if no prices found
                if (unitPrice == "0.00" && totalPrice == "0.00")
                {
                    Console.WriteLine($"⚠️ No prices found in expected columns for {artiCode}, scanning all columns");

                    for (int i = 0; i < columns.Count; i++)
                    {
                        var cellText = columns[i].InnerText?.Trim() ?? "";
                        if (ContainsPricePattern(cellText))
                        {
                            var extractedPrice = ExtractPriceFromString(cellText);
                            if (extractedPrice != "0.00")
                            {
                                if (unitPrice == "0.00")
                                {
                                    unitPrice = extractedPrice;
                                    Console.WriteLine($"💰 Found unit price in column {i}: {extractedPrice}");
                                }
                                else if (totalPrice == "0.00")
                                {
                                    totalPrice = extractedPrice;
                                    Console.WriteLine($"💰 Found total price in column {i}: {extractedPrice}");
                                    break;
                                }
                            }
                        }
                    }
                }

                // Method 3: Calculate missing price if we have one and quantity
                if (unitPrice != "0.00" && totalPrice == "0.00" && columns.Count > 4)
                {
                    var quantityText = columns[4].InnerText?.Trim() ?? "";
                    if (decimal.TryParse(quantityText, out decimal qty) && qty > 0 &&
                        decimal.TryParse(unitPrice, out decimal uPrice))
                    {
                        totalPrice = (uPrice * qty).ToString("F2");
                        Console.WriteLine($"💰 Calculated total price: {unitPrice} x {qty} = {totalPrice}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error extracting prices for {artiCode}: {ex.Message}");
            }

            return (unitPrice, totalPrice);
        }

        /// <summary>
        /// Helper method to extract field values
        /// </summary>
        private string ExtractFieldValue(string dataString, string fieldName)
        {
            try
            {
                int index = dataString.IndexOf(fieldName);
                if (index == -1) return "";

                dataString = dataString.Substring(index);
                string value = dataString.Substring(dataString.IndexOf(":"), dataString.IndexOf("<") - dataString.IndexOf(":"));
                return value.Substring(1);
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Enhanced price extraction from string
        /// </summary>
        private string ExtractPriceFromString(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "0.00";

            try
            {
                var pricePatterns = new[]
                {
            @"€\s*(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})",
            @"\$\s*(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})",
            @"(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s*€",
            @"(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})\s*\$",
            @"(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})"
        };

                foreach (var pattern in pricePatterns)
                {
                    var matches = Regex.Matches(text, pattern);
                    foreach (Match match in matches)
                    {
                        var priceStr = match.Groups[1].Value;
                        var cleanPrice = CleanAndValidatePrice(priceStr);
                        if (cleanPrice != "0.00")
                        {
                            return cleanPrice;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error extracting price from '{text}': {ex.Message}");
            }

            return "0.00";
        }

        /// <summary>
        /// Check if text contains price patterns
        /// </summary>
        private bool ContainsPricePattern(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;

            var patterns = new[]
            {
        @"€\s*\d+[.,]\d{2}",
        @"\$\s*\d+[.,]\d{2}",
        @"\d+[.,]\d{2}\s*€",
        @"\d+[.,]\d{2}\s*\$",
        @"\d+[.,]\d{2}"
    };

            return patterns.Any(pattern => Regex.IsMatch(text, pattern));
        }

        /// <summary>
        /// Clean and validate price
        /// </summary>
        private string CleanAndValidatePrice(string price)
        {
            if (string.IsNullOrWhiteSpace(price)) return "0.00";

            try
            {
                price = price.Replace("€", "").Replace("$", "").Replace("£", "").Trim();

                // Handle European format
                if (price.Contains(",") && !price.Contains("."))
                {
                    price = price.Replace(",", ".");
                }
                else if (price.Contains(",") && price.Contains("."))
                {
                    var lastComma = price.LastIndexOf(',');
                    var lastDot = price.LastIndexOf('.');
                    if (lastComma > lastDot)
                    {
                        price = price.Replace(".", "").Replace(",", ".");
                    }
                }

                price = Regex.Replace(price, @"[^\d\.]", "");

                if (decimal.TryParse(price, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result) && result > 0)
                {
                    return result.ToString("F2", CultureInfo.InvariantCulture);
                }
            }
            catch
            {
                // Fall through to return 0.00
            }

            return "0.00";
        }
        public static string FormatDate(string dateStr, string currentFormat = "dd/MM/yyyy", string desiredFormat = "dd-MM-yyyy")
        {
            if (string.IsNullOrWhiteSpace(dateStr))
            {
                LogDebug($"Date formatting: input is null/empty");
                return string.Empty;
            }

            try
            {
                DateTime date = DateTime.ParseExact(dateStr.Trim(), currentFormat, CultureInfo.InvariantCulture);
                string formattedDate = date.ToString(desiredFormat);
                LogDebug($"Date formatted: '{dateStr}' -> '{formattedDate}'");
                return formattedDate;
            }
            catch (FormatException ex)
            {
                LogDebug($"Date formatting failed for '{dateStr}': {ex.Message}");
                return string.Empty;
            }
        }

        public (string poNumber, string OrderDate) ExtractPoInfo(string body)
        {
            LogDebug("Extracting PO info from HTML body");

            string poNumber = "null";
            string orderDate = "null";

            if (!body.Contains("po_info"))
            {
                LogDebug("No po_info found in body");
                return (poNumber, orderDate);
            }

            try
            {
                var table = Doc.DocumentNode.SelectSingleNode("//td[@id='po_info']");
                if (table == null)
                {
                    LogDebug("po_info table element not found");
                    return (poNumber, orderDate);
                }

                var rows = table.SelectNodes(".//tr");
                LogDebug($"Found {rows?.Count ?? 0} rows in po_info table");

                if (rows != null && rows.Count >= 2)
                {
                    // Extract PO Number
                    string rawPoNumber = rows[0].InnerText;
                    LogDebug($"Raw PO number text: '{rawPoNumber}'");

                    var words = rawPoNumber.Split(" ");
                    if (words.Length > 2)
                    {
                        int endPoNumberIndex = words[2].IndexOf('&');
                        if (endPoNumberIndex != -1)
                        {
                            poNumber = words[2].Substring(0, endPoNumberIndex);
                            LogDebug($"Extracted PO number: '{poNumber}'");
                        }
                    }

                    // Extract PO Date
                    string rawPoDate = rows[1].InnerText;
                    LogDebug($"Raw PO date text: '{rawPoDate}'");

                    var words2 = rawPoDate.Split(" ");
                    if (words2.Length > 1)
                    {
                        int endPoDateIndex = words2[1].IndexOf("&");
                        if (endPoDateIndex != -1)
                        {
                            orderDate = words2[1].Substring(0, endPoDateIndex);
                            LogDebug($"Extracted PO date: '{orderDate}'");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogDebug($"ERROR extracting PO info: {ex.Message}");
            }

            return (poNumber, orderDate);
        }

        public List<OrderLine> ExtractOrderLinesWeirDomain()
        {
            try
            {
                var orderLines = new List<OrderLine>();

                // Search for the table with id 'order_lines'
                var table = Doc.DocumentNode.SelectSingleNode("//table[@id='order_lines']");
                if (table == null)
                {
                    Console.WriteLine("❌ No order_lines table found");
                    return orderLines;
                }

                // Step 1: Analyze table structure to find column positions
                var columnMapping = AnalyzeTableStructure(table);
                Console.WriteLine($"📊 Detected column mapping: {string.Join(", ", columnMapping.Select(kv => $"{kv.Key}={kv.Value}"))}");

                // Search all rows in the table body (tbody)
                var rows = table.SelectNodes(".//tr");
                foreach (var row in rows)
                {
                    var columns = row.SelectNodes(".//td");
                    if (columns == null)
                    {
                        continue; // Skip the row if it doesn't contain <td> elements
                    }

                    // Skip header rows or insufficient columns
                    if (columns.Count < Math.Max(columnMapping.GetValueOrDefault("UnitPrice", 6),
                                                columnMapping.GetValueOrDefault("TotalPrice", 7)))
                    {
                        continue;
                    }

                    string dataString = columns[2].InnerHtml;

                    // Skip header rows
                    if (dataString.Contains(">Description </span>"))
                    {
                        continue;
                    }

                    try
                    {
                        // Extract basic data (your existing logic)
                        var basicData = ExtractBasicOrderData(dataString, columns);
                        if (basicData.artiCode.Contains("SUBCON:"))
                        {
                            continue;
                        }

                        // Enhanced price extraction using dynamic column mapping
                        var priceData = ExtractPricesWithDynamicMapping(columns, columnMapping, basicData.artiCode);

                        OrderLine order = new OrderLine
                        (
                            FilterInput(columns[0].InnerText.Trim()),
                            FilterInput(basicData.artiCode),
                            FilterInput(basicData.description),
                            FilterInput(basicData.drawingNumber),
                            FilterInput(basicData.revision),
                            FilterInput(basicData.supplierPartNumber),
                            "not implemented",
                            FilterInput(basicData.requestedShippingDate),
                            FilterInput(basicData.articleDescription),
                            FormatDate(columns[3].InnerText.Trim()),
                            FilterInput(basicData.sapPoLineNumber),
                            FilterInput(columns[4].InnerText.Trim()),
                            FilterInput(columns[5].InnerText.Trim()),
                            FilterInput(priceData.unitPrice),    // Enhanced price extraction
                            FilterInput(priceData.totalPrice)    // Enhanced price extraction
                        );

                        orderLines.Add(order);
                        Console.WriteLine($"✅ Extracted order: {basicData.artiCode} - Unit: {priceData.unitPrice}, Total: {priceData.totalPrice}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error processing row: {ex.Message}");
                        continue;
                    }
                }

                Console.WriteLine($"📦 Total orders extracted: {orderLines.Count}");
                return orderLines;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in enhanced HTML parser: {ex.Message}");
                return new List<OrderLine>();
            }
        }

        public List<OrderLine> ExtractOrderLinesOutlookForwardMail()
        {
            LogDebug("=== STARTING OUTLOOK FORWARD MAIL EXTRACTION ===");
            ClearDebugLog();

            try
            {
                var orderLines = new List<OrderLine>();

                if (Doc == null)
                {
                    LogDebug("ERROR: Document is null");
                    return orderLines;
                }

                var table = Doc.DocumentNode.SelectSingleNode("//table[@id='order_lines']");
                if (table == null)
                {
                    LogDebug("ERROR: No table with id 'order_lines' found for Outlook forward mail");
                    return orderLines;
                }

                LogDebug("✓ Found order_lines table for Outlook forward mail");

                var rows = table.SelectNodes(".//tr");
                LogDebug($"Found {rows?.Count ?? 0} rows in order table");

                if (rows == null) return orderLines;

                int rowIndex = 0;
                foreach (var row in rows)
                {
                    rowIndex++;
                    LogDebug($"\n--- Processing Outlook Row {rowIndex} ---");

                    var columns = row.SelectNodes(".//td");
                    if (columns == null)
                    {
                        LogDebug($"Row {rowIndex}: No <td> elements found");
                        continue;
                    }

                    if (columns.Count < 7)
                    {
                        LogDebug($"Row {rowIndex}: Not enough columns ({columns.Count} < 7)");
                        continue;
                    }

                    try
                    {
                        string dataString = columns[2].InnerHtml;
                        LogDebug($"Row {rowIndex}: Processing column 2 HTML (length: {dataString?.Length ?? 0})");

                        // Skip description header
                        if (dataString.Contains(">Description </span>"))
                        {
                            LogDebug($"Row {rowIndex}: Header row, skipping");
                            continue;
                        }

                        // Extract description for Outlook forward format
                        dataString = dataString.Substring(90); // Skip initial formatting
                        dataString = dataString.Substring(dataString.IndexOf(">") + 1);
                        int descriptionEndIndex = dataString.IndexOf("</span>");

                        if (descriptionEndIndex == -1)
                        {
                            LogDebug($"Row {rowIndex}: No description end tag found");
                            continue;
                        }

                        string description = dataString.Substring(0, descriptionEndIndex);
                        if (string.IsNullOrEmpty(description) || description.Contains("Description"))
                        {
                            LogDebug($"Row {rowIndex}: Invalid description, skipping");
                            continue;
                        }

                        LogDebug($"Row {rowIndex}: Description: '{description}'");

                        string artiCode = description.Contains(" ") ?
                            description.Substring(0, description.IndexOf(" ")) :
                            description;

                        if (artiCode.Contains('A'))
                        {
                            artiCode = artiCode.Substring(0, artiCode.IndexOf('A'));
                        }

                        // Extract drawing number and revision
                        int drawingNumberIndex = dataString.IndexOf("Drawing Number:");
                        if (drawingNumberIndex == -1)
                        {
                            LogDebug($"Row {rowIndex}: No drawing number found");
                            continue;
                        }

                        dataString = dataString.Substring(drawingNumberIndex);
                        string drawingNumberRaw = dataString.Substring(
                            dataString.IndexOf(":"),
                            dataString.IndexOf("<") + 1 - dataString.IndexOf(":")
                        );

                        string drawingNumber = drawingNumberRaw.Substring(1, drawingNumberRaw.IndexOf("rev.") - 1).Trim();
                        string revisionRaw = drawingNumberRaw.Substring(drawingNumberRaw.IndexOf("rev."));
                        string revision = revisionRaw.Substring(revisionRaw.IndexOf(" "), 2).Trim();

                        LogDebug($"Row {rowIndex}: Drawing: '{drawingNumber}', Revision: '{revision}'");

                        // Extract other fields
                        string supplierPartNumber = ExtractFieldValue(dataString, "Supplier Part Number:");
                        string requestedShippingDate = ExtractFieldValue(dataString, "Requested Ship Date:");
                        string sapPoLineNumber = ExtractFieldValue(dataString, "SAP PO Line Number:");

                        // Extract article description for Outlook format
                        string articleDescription = ExtractOutlookArticleDescription(dataString);

                        if (artiCode.Contains("SUBCON:"))
                        {
                            LogDebug($"Row {rowIndex}: SUBCON entry, skipping");
                            continue;
                        }

                        OrderLine order = new OrderLine(
                            FilterInput(columns[0].InnerText.Trim()),
                            FilterInput(artiCode),
                            FilterInput(description),
                            FilterInput(drawingNumber),
                            FilterInput(revision),
                            FilterInput(supplierPartNumber),
                            "not implemented",
                            FilterInput(requestedShippingDate),
                            FilterInput(articleDescription),
                            FormatDate(columns[3].InnerText.Trim()),
                            FilterInput(sapPoLineNumber),
                            FilterInput(columns[4].InnerText.Trim()),
                            FilterInput(columns[5].InnerText.Trim()),
                            FilterInput(columns[6].InnerText.Trim()),
                            FilterInput(columns[7].InnerText.Trim())
                        );

                        order.SetExtractionDetails("HTML_ELCOTEC_FORWARD", $"Outlook Row {rowIndex}, Article: {artiCode}");
                        orderLines.Add(order);
                        LogDebug($"Row {rowIndex}: ✓ Outlook order line created successfully");

                    }
                    catch (Exception rowEx)
                    {
                        LogDebug($"Row {rowIndex}: ERROR processing Outlook row: {rowEx.Message}");
                        continue;
                    }
                }

                LogDebug($"\n=== OUTLOOK EXTRACTION COMPLETE ===");
                LogDebug($"Successfully extracted {orderLines.Count} order lines");

                return orderLines;
            }
            catch (Exception ex)
            {
                LogDebug($"CRITICAL ERROR in ExtractOrderLinesOutlookForwardMail: {ex.Message}");
                return new List<OrderLine>();
            }
        }
     
        private string ExtractArticleDescription(string dataString, string requestedShippingDate, string sapPoLineNumber)
        {
            try
            {
                int requestShipDateIndex = dataString.IndexOf("Requested Ship Date:");
                int sapPoLineNumberIndex = dataString.IndexOf("SAP PO Line Number:");

                bool descriptionAvailable = sapPoLineNumberIndex - requestShipDateIndex > 160;

                if (!descriptionAvailable) return "";

                string extractionString = dataString.Substring(requestShipDateIndex + 50);

                if (extractionString.Length < 900)
                {
                    int startIndex = extractionString.IndexOf("width") + 12;
                    int endIndex = extractionString.IndexOf("<br>");
                    if (startIndex > 11 && endIndex > startIndex)
                    {
                        return extractionString.Substring(startIndex, endIndex - startIndex);
                    }
                }
                else
                {
                    extractionString = extractionString.Substring(extractionString.IndexOf("Tracking Number:") + 50);
                    int startIndex = extractionString.IndexOf("width") + 12;
                    int endIndex = extractionString.IndexOf("<br>");
                    if (startIndex > 11 && endIndex > startIndex)
                    {
                        return extractionString.Substring(startIndex, endIndex - startIndex);
                    }
                }

                return "";
            }
            catch
            {
                return "";
            }
        }

        private string ExtractOutlookArticleDescription(string dataString)
        {
            try
            {
                if (dataString.Contains(">Tracking Number: "))
                {
                    string trackingSection = dataString.Substring(dataString.IndexOf(">Tracking Number: "));
                    trackingSection = trackingSection.Substring(150);
                    int startIndex = trackingSection.IndexOf(">");
                    int endIndex = trackingSection.IndexOf("<br>");

                    if (startIndex != -1 && endIndex > startIndex)
                    {
                        return trackingSection.Substring(startIndex + 1, endIndex - startIndex - 1);
                    }
                }

                return "";
            }
            catch
            {
                return "";
            }
        }

        public string FilterInput(string input)
        {
            if (input == null) return "ERROR!";

            return input.Contains("\r\n") ? input.Replace("\r\n", "") : input;
        }
    }
}