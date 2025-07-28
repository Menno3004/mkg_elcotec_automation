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
    class QuoteHtmlParser
    {
        public QuoteHtmlParser() { }

        public static HtmlDocument Doc { get; set; }

        private static List<string> DebugLog = new List<string>();

        public static void ClearDebugLog() => DebugLog.Clear();
        public static List<string> GetDebugLog() => new List<string>(DebugLog);

        private static void LogDebug(string message)
        {
            var logEntry = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
            DebugLog.Add(logEntry);
            Console.WriteLine($"[QUOTE_HTML_PARSER] {logEntry}");
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

                // Check for quote_lines table
                var quoteTable = doc.DocumentNode?.SelectSingleNode("//table[@id='quote_lines']");
                LogDebug($"Quote lines table found: {quoteTable != null}");

                if (quoteTable != null)
                {
                    var rows = quoteTable.SelectNodes(".//tr");
                    LogDebug($"Quote table has {rows?.Count ?? 0} rows");
                }

                // Check for RFQ tables
                var rfqTable = doc.DocumentNode?.SelectSingleNode("//table[@id='rfq_lines']");
                LogDebug($"RFQ lines table found: {rfqTable != null}");

                // Check for other potential quote tables
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

        public (string rfqNumber, string quoteDate, string validUntil) ExtractRfqInfo(string body)
        {
            LogDebug("Extracting RFQ info from HTML body");

            string rfqNumber = "null";
            string quoteDate = "null";
            string validUntil = "null";

            if (!body.Contains("rfq_info") && !body.Contains("quote_info"))
            {
                LogDebug("No rfq_info or quote_info found in body");
                return (rfqNumber, quoteDate, validUntil);
            }

            try
            {
                var table = Doc.DocumentNode.SelectSingleNode("//td[@id='rfq_info']") ??
                           Doc.DocumentNode.SelectSingleNode("//td[@id='quote_info']");
                if (table == null)
                {
                    LogDebug("rfq_info/quote_info table element not found");
                    return (rfqNumber, quoteDate, validUntil);
                }

                var rows = table.SelectNodes(".//tr");
                LogDebug($"Found {rows?.Count ?? 0} rows in rfq_info table");

                if (rows != null && rows.Count >= 3)
                {
                    // Extract RFQ Number
                    string rawRfqNumber = rows[0].InnerText;
                    LogDebug($"Raw RFQ number text: '{rawRfqNumber}'");

                    var words = rawRfqNumber.Split(" ");
                    if (words.Length > 2)
                    {
                        int endRfqNumberIndex = words[2].IndexOf('&');
                        if (endRfqNumberIndex != -1)
                        {
                            rfqNumber = words[2].Substring(0, endRfqNumberIndex);
                            LogDebug($"Extracted RFQ number: '{rfqNumber}'");
                        }
                        else
                        {
                            rfqNumber = words[2];
                            LogDebug($"Extracted RFQ number (no delimiter): '{rfqNumber}'");
                        }
                    }

                    // Extract Quote Date
                    string rawQuoteDate = rows[1].InnerText;
                    LogDebug($"Raw quote date text: '{rawQuoteDate}'");

                    var words2 = rawQuoteDate.Split(" ");
                    if (words2.Length > 1)
                    {
                        int endQuoteDateIndex = words2[1].IndexOf("&");
                        if (endQuoteDateIndex != -1)
                        {
                            quoteDate = words2[1].Substring(0, endQuoteDateIndex);
                            LogDebug($"Extracted quote date: '{quoteDate}'");
                        }
                        else
                        {
                            quoteDate = words2[1];
                            LogDebug($"Extracted quote date (no delimiter): '{quoteDate}'");
                        }
                    }

                    // Extract Valid Until (if available)
                    if (rows.Count > 2)
                    {
                        string rawValidUntil = rows[2].InnerText;
                        LogDebug($"Raw valid until text: '{rawValidUntil}'");

                        var words3 = rawValidUntil.Split(" ");
                        if (words3.Length > 1)
                        {
                            int endValidUntilIndex = words3[1].IndexOf("&");
                            if (endValidUntilIndex != -1)
                            {
                                validUntil = words3[1].Substring(0, endValidUntilIndex);
                                LogDebug($"Extracted valid until: '{validUntil}'");
                            }
                            else
                            {
                                validUntil = words3[1];
                                LogDebug($"Extracted valid until (no delimiter): '{validUntil}'");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogDebug($"ERROR extracting RFQ info: {ex.Message}");
            }

            return (rfqNumber, quoteDate, validUntil);
        }

        public List<QuoteLine> ExtractQuoteLinesWeirDomain()
        {
            LogDebug("=== STARTING WEIR DOMAIN QUOTE EXTRACTION ===");
            ClearDebugLog();

            try
            {
                var quoteLines = new List<QuoteLine>();

                if (Doc == null)
                {
                    LogDebug("ERROR: Document is null");
                    return quoteLines;
                }

                // Search for the table with id 'quote_lines' or 'rfq_lines'
                var table = Doc.DocumentNode.SelectSingleNode("//table[@id='quote_lines']") ??
                           Doc.DocumentNode.SelectSingleNode("//table[@id='rfq_lines']");

                if (table == null)
                {
                    LogDebug("ERROR: No table with id 'quote_lines' or 'rfq_lines' found");

                    // Try alternative selectors
                    var allTables = Doc.DocumentNode.SelectNodes("//table");
                    LogDebug($"Found {allTables?.Count ?? 0} tables in document");

                    if (allTables != null)
                    {
                        foreach (var t in allTables)
                        {
                            var tableId = t.GetAttributeValue("id", "");
                            var tableClass = t.GetAttributeValue("class", "");
                            LogDebug($"Table found: id='{tableId}', class='{tableClass}'");

                            // Look for tables that might contain quote data
                            if (tableId.ToLower().Contains("quote") || tableId.ToLower().Contains("rfq") ||
                                tableClass.ToLower().Contains("quote") || tableClass.ToLower().Contains("rfq"))
                            {
                                table = t;
                                LogDebug($"Using alternative table: id='{tableId}', class='{tableClass}'");
                                break;
                            }
                        }
                    }

                    if (table == null)
                    {
                        LogDebug("No suitable quote table found");
                        return quoteLines;
                    }
                }

                LogDebug("✓ Found quote lines table for Weir domain");

                var rows = table.SelectNodes(".//tr");
                LogDebug($"Found {rows?.Count ?? 0} rows in quote table");

                if (rows == null) return quoteLines;

                // Extract RFQ info once for all lines
                var rfqInfo = ExtractRfqInfo(Doc.DocumentNode.InnerHtml);
                LogDebug($"RFQ Info: Number='{rfqInfo.rfqNumber}', Date='{rfqInfo.quoteDate}', ValidUntil='{rfqInfo.validUntil}'");

                int rowIndex = 0;
                foreach (var row in rows)
                {
                    rowIndex++;
                    LogDebug($"\n--- Processing Weir Quote Row {rowIndex} ---");

                    var columns = row.SelectNodes(".//td");
                    if (columns == null)
                    {
                        LogDebug($"Row {rowIndex}: No <td> elements found");
                        continue;
                    }

                    if (columns.Count < 6) // Minimum columns for quote line
                    {
                        LogDebug($"Row {rowIndex}: Not enough columns ({columns.Count} < 6)");
                        continue;
                    }

                    try
                    {
                        // Column structure for quote lines (adjust based on actual structure):
                        // 0: Line Number
                        // 1: Article Code
                        // 2: Description
                        // 3: Quantity
                        // 4: Unit
                        // 5: Quoted Price
                        // 6: Customer Part Number (if available)
                        // 7: Drawing Number (if available)
                        // 8: Revision (if available)

                        string lineNumber = columns[0].InnerText?.Trim() ?? rowIndex.ToString();
                        string artiCode = columns[1].InnerText?.Trim() ?? "";
                        string description = columns[2].InnerText?.Trim() ?? "";
                        string quantity = columns[3].InnerText?.Trim() ?? "1";
                        string unit = columns[4].InnerText?.Trim() ?? "PCS";
                        string quotedPrice = columns[5].InnerText?.Trim() ?? "0.00";

                        // Optional columns
                        string customerPartNumber = columns.Count > 6 ? columns[6].InnerText?.Trim() ?? "" : "";
                        string drawingNumber = columns.Count > 7 ? columns[7].InnerText?.Trim() ?? "" : "";
                        string revision = columns.Count > 8 ? columns[8].InnerText?.Trim() ?? "00" : "00";

                        LogDebug($"Row {rowIndex}: Raw data - Line: '{lineNumber}', ArtiCode: '{artiCode}', Desc: '{description}', Qty: '{quantity}', Unit: '{unit}', Price: '{quotedPrice}'");

                        // Skip header rows or empty rows
                        if (string.IsNullOrWhiteSpace(artiCode) ||
                            artiCode.ToLower().Contains("article") ||
                            artiCode.ToLower().Contains("code") ||
                            artiCode.ToLower().Contains("line") ||
                            description.ToLower().Contains("description"))
                        {
                            LogDebug($"Row {rowIndex}: Skipping header or empty row");
                            continue;
                        }

                        // Clean price data
                        quotedPrice = CleanPrice(quotedPrice);
                        quantity = CleanQuantity(quantity);

                        var quote = new QuoteLine(
                            lineNumber: lineNumber,
                            artiCode: artiCode,
                            description: description,
                            quantity: quantity,
                            unit: unit,
                            quotedPrice: quotedPrice,
                            customerPartNumber: customerPartNumber,
                            rfqNumber: rfqInfo.rfqNumber != "null" ? rfqInfo.rfqNumber : "",
                            drawingNumber: drawingNumber,
                            revision: revision,
                            requestedDeliveryDate: "",
                            quoteDate: rfqInfo.quoteDate != "null" ? rfqInfo.quoteDate : DateTime.Now.ToString("dd-MM-yyyy"),
                            quoteStatus: "Draft",
                            priority: DeterminePriority(quotedPrice),
                            validUntil: rfqInfo.validUntil != "null" ? rfqInfo.validUntil : DateTime.Now.AddDays(30).ToString("dd-MM-yyyy"),
                            extractionMethod: "WEIR_DOMAIN_HTML",
                            extractionDomain: "weir.com"
                        );

                        LogDebug($"Row {rowIndex}: ✓ Quote line created - RFQ: {quote.RfqNumber}, Article: {artiCode}, Price: {quotedPrice}");

                        quoteLines.Add(quote);
                        LogDebug($"Row {rowIndex}: ✓ Quote line added successfully");

                    }
                    catch (Exception rowEx)
                    {
                        LogDebug($"Row {rowIndex}: ERROR processing row: {rowEx.Message}");
                        continue; // Continue with next row
                    }
                }

                LogDebug($"\n=== WEIR QUOTE EXTRACTION COMPLETE ===");
                LogDebug($"Successfully extracted {quoteLines.Count} quote lines");

                return quoteLines;
            }
            catch (Exception ex)
            {
                LogDebug($"CRITICAL ERROR in ExtractQuoteLinesWeirDomain: {ex.Message}");
                LogDebug($"Stack trace: {ex.StackTrace}");
                return new List<QuoteLine>();
            }
        }

        public List<QuoteLine> ExtractQuoteLinesOutlookForwardMail()
        {
            LogDebug("=== STARTING OUTLOOK FORWARD MAIL QUOTE EXTRACTION ===");
            ClearDebugLog();

            try
            {
                var quoteLines = new List<QuoteLine>();

                if (Doc == null)
                {
                    LogDebug("ERROR: Document is null");
                    return quoteLines;
                }

                var table = Doc.DocumentNode.SelectSingleNode("//table[@id='quote_lines']") ??
                           Doc.DocumentNode.SelectSingleNode("//table[@id='rfq_lines']");

                if (table == null)
                {
                    LogDebug("ERROR: No table with id 'quote_lines' or 'rfq_lines' found for Outlook forward mail");
                    return quoteLines;
                }

                LogDebug("✓ Found quote lines table for Outlook forward mail");

                var rows = table.SelectNodes(".//tr");
                LogDebug($"Found {rows?.Count ?? 0} rows in quote table");

                if (rows == null) return quoteLines;

                // Extract RFQ info once for all lines
                var rfqInfo = ExtractRfqInfo(Doc.DocumentNode.InnerHtml);
                LogDebug($"RFQ Info: Number='{rfqInfo.rfqNumber}', Date='{rfqInfo.quoteDate}', ValidUntil='{rfqInfo.validUntil}'");

                int rowIndex = 0;
                foreach (var row in rows)
                {
                    rowIndex++;
                    LogDebug($"\n--- Processing Outlook Quote Row {rowIndex} ---");

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
                        if (dataString.Contains(">Description </span>") || dataString.Contains(">RFQ Item </span>"))
                        {
                            LogDebug($"Row {rowIndex}: Header row, skipping");
                            continue;
                        }

                        // Extract description for Outlook forward format (similar to order processing)
                        dataString = dataString.Substring(Math.Min(90, dataString.Length)); // Skip initial formatting
                        int firstGtIndex = dataString.IndexOf(">");
                        if (firstGtIndex != -1)
                        {
                            dataString = dataString.Substring(firstGtIndex + 1);
                        }

                        int descriptionEndIndex = dataString.IndexOf("</span>");
                        if (descriptionEndIndex == -1)
                        {
                            LogDebug($"Row {rowIndex}: No description end tag found");
                            continue;
                        }

                        string description = dataString.Substring(0, descriptionEndIndex);
                        if (string.IsNullOrEmpty(description) || description.Contains("Description") || description.Contains("RFQ Item"))
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

                        // Extract drawing number and revision (if available)
                        string drawingNumber = "";
                        string revision = "00";

                        int drawingNumberIndex = dataString.IndexOf("Drawing Number:");
                        if (drawingNumberIndex != -1)
                        {
                            dataString = dataString.Substring(drawingNumberIndex);
                            string drawingNumberRaw = dataString.Substring(
                                dataString.IndexOf(":"),
                                Math.Min(dataString.IndexOf("<") + 1 - dataString.IndexOf(":"), dataString.Length - dataString.IndexOf(":"))
                            );

                            if (drawingNumberRaw.Contains("rev."))
                            {
                                drawingNumber = drawingNumberRaw.Substring(1, drawingNumberRaw.IndexOf("rev.") - 1).Trim();
                                string revisionRaw = drawingNumberRaw.Substring(drawingNumberRaw.IndexOf("rev.") + 4);
                                revision = revisionRaw.Substring(0, Math.Min(revisionRaw.IndexOf("<"), revisionRaw.Length)).Trim();
                            }
                            else
                            {
                                drawingNumber = drawingNumberRaw.Substring(1).Replace("<", "").Trim();
                            }
                        }

                        // Extract quantities and prices from other columns
                        string quantity = ExtractColumnText(columns[3]);
                        string unit = ExtractColumnText(columns[4]) ?? "PCS";
                        string quotedPrice = ExtractColumnText(columns[5]);

                        // Clean the extracted data
                        quantity = CleanQuantity(quantity);
                        quotedPrice = CleanPrice(quotedPrice);

                        var quote = new QuoteLine(
                            lineNumber: rowIndex.ToString(),
                            artiCode: artiCode,
                            description: description,
                            quantity: quantity,
                            unit: unit,
                            quotedPrice: quotedPrice,
                            customerPartNumber: artiCode, // Use artiCode as customer part number for now
                            rfqNumber: rfqInfo.rfqNumber != "null" ? rfqInfo.rfqNumber : $"RFQ-{DateTime.Now:yyyyMMdd}-{rowIndex:D3}",
                            drawingNumber: drawingNumber,
                            revision: revision,
                            requestedDeliveryDate: "",
                            quoteDate: rfqInfo.quoteDate != "null" ? rfqInfo.quoteDate : DateTime.Now.ToString("dd-MM-yyyy"),
                            quoteStatus: "Draft",
                            priority: DeterminePriority(quotedPrice),
                            validUntil: rfqInfo.validUntil != "null" ? rfqInfo.validUntil : DateTime.Now.AddDays(30).ToString("dd-MM-yyyy"),
                            extractionMethod: "OUTLOOK_FORWARD_MAIL_HTML",
                            extractionDomain: "outlook.com"
                        );

                        LogDebug($"Row {rowIndex}: ✓ Quote line created - RFQ: {quote.RfqNumber}, Article: {artiCode}");

                        quoteLines.Add(quote);
                        LogDebug($"Row {rowIndex}: ✓ Quote line added successfully");

                    }
                    catch (Exception rowEx)
                    {
                        LogDebug($"Row {rowIndex}: ERROR processing row: {rowEx.Message}");
                        continue; // Continue with next row
                    }
                }

                LogDebug($"\n=== OUTLOOK QUOTE EXTRACTION COMPLETE ===");
                LogDebug($"Successfully extracted {quoteLines.Count} quote lines");

                return quoteLines;
            }
            catch (Exception ex)
            {
                LogDebug($"CRITICAL ERROR in ExtractQuoteLinesOutlookForwardMail: {ex.Message}");
                LogDebug($"Stack trace: {ex.StackTrace}");
                return new List<QuoteLine>();
            }
        }

        #region Helper Methods

        private static string ExtractColumnText(HtmlNode column)
        {
            if (column == null) return "";

            // Try to get clean text content
            string text = column.InnerText?.Trim() ?? "";

            // If text is empty, try to extract from HTML
            if (string.IsNullOrWhiteSpace(text))
            {
                string html = column.InnerHtml;
                if (!string.IsNullOrWhiteSpace(html))
                {
                    // Remove HTML tags and get text
                    text = System.Text.RegularExpressions.Regex.Replace(html, "<.*?>", "").Trim();
                }
            }

            return text;
        }

        private static string CleanPrice(string price)
        {
            if (string.IsNullOrWhiteSpace(price))
                return "0.00";

            // Remove currency symbols and extra whitespace
            price = price.Replace("€", "").Replace("$", "").Replace("£", "")
                        .Replace("USD", "").Replace("EUR", "").Replace("GBP", "")
                        .Trim();

            // Handle comma as decimal separator
            if (price.Contains(",") && !price.Contains("."))
            {
                price = price.Replace(",", ".");
            }
            else if (price.Contains(",") && price.Contains("."))
            {
                // European format: 1.234,56 -> 1234.56
                var lastComma = price.LastIndexOf(',');
                var lastDot = price.LastIndexOf('.');
                if (lastComma > lastDot)
                {
                    price = price.Replace(".", "").Replace(",", ".");
                }
            }

            // Remove any non-numeric characters except decimal point
            price = System.Text.RegularExpressions.Regex.Replace(price, @"[^\d\.]", "");

            // Validate result
            if (decimal.TryParse(price, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
            {
                return result.ToString("F2", CultureInfo.InvariantCulture);
            }

            return "0.00";
        }

        private static string CleanQuantity(string quantity)
        {
            if (string.IsNullOrWhiteSpace(quantity))
                return "1";

            // Remove units and extra text
            quantity = quantity.Replace("PCS", "").Replace("EA", "").Replace("PC", "")
                              .Replace("EACH", "").Replace("ST", "").Replace("STK", "")
                              .Trim();

            // Remove any non-numeric characters except decimal point
            quantity = System.Text.RegularExpressions.Regex.Replace(quantity, @"[^\d\.]", "");

            // Validate result
            if (decimal.TryParse(quantity, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result) && result > 0)
            {
                return result.ToString("F0", CultureInfo.InvariantCulture);
            }

            return "1";
        }

        private static string DeterminePriority(string quotedPrice)
        {
            if (decimal.TryParse(quotedPrice?.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal price))
            {
                if (price > 1000)
                    return "High";
                else if (price > 100)
                    return "Medium";
                else
                    return "Normal";
            }
            return "Normal";
        }

        #endregion
    }
}