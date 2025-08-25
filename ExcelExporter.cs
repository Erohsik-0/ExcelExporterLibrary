using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelExport
{
    public class ExcelExporter
    {
        // Cache for parsed data types to avoid repeated parsing attempts
        private readonly Dictionary<string, Type> _typeCache = new Dictionary<string, Type>();

        /// <summary>
        /// Exports a list of dictionaries to Excel format with proper styling and data types
        /// </summary>
        /// <param name="dataList">The data to export as key-value pairs</param>
        /// <param name="sheetName">Name of the worksheet (defaults to "Data")</param>
        /// <returns>Excel file as byte array</returns>
        public byte[] ExportToExcel(List<Dictionary<string, object>> dataList, string sheetName = "Data")
        {
            // Handle empty data gracefully - create a simple placeholder sheet
            if (dataList == null || dataList.Count == 0)
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(sheetName);
                    worksheet.Cell(1, 1).Value = "No data available";
                    using var stream = new MemoryStream();
                    workbook.SaveAs(stream);
                    return stream.ToArray();
                }
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(sheetName);

                // Extract headers from the first row to maintain consistent column order
                var headers = dataList[0].Keys.ToArray(); // Use array for better performance
                var headerCount = headers.Length;

                // Create header row with styling in batch
                var headerRange = worksheet.Range(1, 1, 1, headerCount);
                for (int col = 0; col < headerCount; col++)
                {
                    worksheet.Cell(1, col + 1).Value = headers[col];
                }
                StyleHeaderRange(headerRange);

                // Set up filtering and freeze the header row for better user experience
                headerRange.SetAutoFilter();
                worksheet.SheetView.FreezeRows(1);

                // Bulk insert data rows for better performance
                var dataRowCount = dataList.Count;
                for (int row = 0; row < dataRowCount; row++)
                {
                    var rowData = dataList[row];
                    for (int col = 0; col < headerCount; col++)
                    {
                        var key = headers[col];
                        if (rowData.TryGetValue(key, out var value))
                        {
                            var cell = worksheet.Cell(row + 2, col + 1);
                            SetCellValueWithType(cell, value);
                        }
                    }
                }

                // Apply data styling to all data cells at once
                if (dataRowCount > 0)
                {
                    var dataRange = worksheet.Range(2, 1, dataRowCount + 1, headerCount);
                    StyleDataRange(dataRange);
                }

                // Auto-adjust column widths for better readability
                worksheet.Columns().AdjustToContents();

                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                return stream.ToArray();
            }
        }

        /// <summary>
        /// Converts JSON string to Excel by deserializing to dictionary list
        /// </summary>
        public byte[] ExportJsonToExcel(string jsonString)
        {
            var dataList = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(jsonString);
            return ExportToExcel(dataList);
        }

        /// <summary>
        /// Creates multiple Excel sheets grouped by nested content structure
        /// Enhanced to properly handle Cosmos DB response structures
        /// </summary>
        /// <param name="jsonString">JSON data to process</param>
        /// <param name="groupByField">Optional field to use for additional grouping logic</param>
        public byte[] ExportGroupedJsonToExcel(string jsonString, string groupByField = "category")
        {
            try
            {
                // Parse JSON while preserving structure information
                var jsonToken = JToken.Parse(jsonString);
                var structuredData = new List<StructuredRecord>();

                // Extract records while maintaining structure information
                ExtractRecordsWithStructure(jsonToken, structuredData);

                if (structuredData.Count == 0)
                {
                    return ExportToExcel(new List<Dictionary<string, object>>(), "Data");
                }

                using var workbook = new XLWorkbook();

                // Group records by their nested content structure
                var structureGroups = GroupRecordsByStructure(structuredData);

                if (structureGroups.Count <= 1)
                {
                    // No meaningful grouping found, create single sheet
                    var allData = structuredData.Select(r => r.FlatData).ToList();
                    var headers = GetAllUniqueHeaders(allData);
                    CreateWorksheet(workbook, "Data", allData, headers);
                }
                else
                {
                    // Create sheets for each structure group
                    foreach (var group in structureGroups.OrderBy(g => g.Key))
                    {
                        var sanitizedName = SanitizeSheetName(group.Key);
                        var groupData = group.Value.Select(r => r.FlatData).ToList();
                        var headers = GetAllUniqueHeaders(groupData);
                        CreateWorksheet(workbook, sanitizedName, groupData, headers);
                    }

                    // Create comprehensive summary sheet
                    CreateEnhancedSummarySheet(workbook, structureGroups);
                }

                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                return stream.ToArray();
            }
            catch (JsonReaderException ex)
            {
                // Return error sheet for invalid JSON
                return CreateErrorSheet($"Invalid JSON: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Return error sheet for any other issues
                return CreateErrorSheet($"Processing error: {ex.Message}");
            }
        }

        /// <summary>
        /// Flattens nested JSON structures into a single-level Excel sheet
        /// Converts objects like {user: {name: "John", age: 30}} into columns like "user.name", "user.age"
        /// </summary>
        public byte[] ExportFlattenedJsonToExcel(string jsonString)
        {
            var flatList = new List<Dictionary<string, object>>();
            var orderedHeaders = new List<string>();

            try
            {
                var token = JToken.Parse(jsonString);
                bool isFirstItem = true;

                // Handle different JSON root structures
                if (token is JObject obj)
                {
                    // Look for common array containers
                    if (obj["Documents"] is JArray docs)
                    {
                        ProcessJsonArray(docs, flatList, ref orderedHeaders, ref isFirstItem);
                    }
                    else if (obj["Items"] is JArray items)
                    {
                        ProcessJsonArray(items, flatList, ref orderedHeaders, ref isFirstItem);
                    }
                    else if (obj["Data"] is JArray data)
                    {
                        ProcessJsonArray(data, flatList, ref orderedHeaders, ref isFirstItem);
                    }
                    else if (obj["Records"] is JArray records)
                    {
                        ProcessJsonArray(records, flatList, ref orderedHeaders, ref isFirstItem);
                    }
                    else
                    {
                        // Single object case
                        flatList.Add(FlattenJson(obj, "", isFirstItem, ref orderedHeaders));
                    }
                }
                else if (token is JArray arr)
                {
                    ProcessJsonArray(arr, flatList, ref orderedHeaders, ref isFirstItem);
                }
            }
            catch (JsonReaderException)
            {
                // Return empty array for invalid JSON instead of throwing
                return Array.Empty<byte>();
            }

            if (flatList.Count == 0)
            {
                return Array.Empty<byte>();
            }

            return CreateFlattenedExcelWorkbook(flatList, orderedHeaders);
        }

        /// <summary>
        /// Combines flattening with grouping by nested content structure
        /// Groups flattened records based on their original nested content patterns
        /// </summary>
        public byte[] ExportFlattenedGroupedJsonToExcel(string jsonString, string groupByField = "type")
        {
            List<Dictionary<string, object>> flatList = new();
            List<string> orderedHeaders = new();
            List<Dictionary<string, object>> originalStructures = new();

            try
            {
                var token = JToken.Parse(jsonString);
                bool isFirstItem = true;

                if (token is JObject obj)
                {
                    // Check for various common container patterns
                    var arrayContainers = new[] { "Documents", "Items", "Data", "Records", "Results", "Entities" };
                    JArray targetArray = null;

                    foreach (var container in arrayContainers)
                    {
                        if (obj[container] is JArray arr)
                        {
                            targetArray = arr;
                            break;
                        }
                    }

                    if (targetArray != null)
                    {
                        ProcessJsonArrayWithStructure(targetArray, flatList, originalStructures, ref orderedHeaders, ref isFirstItem);
                    }
                    else
                    {
                        flatList.Add(FlattenJson(obj, "", isFirstItem, ref orderedHeaders));
                        originalStructures.Add(ExtractStructureInfo(obj));
                    }
                }
                else if (token is JArray arr)
                {
                    ProcessJsonArrayWithStructure(arr, flatList, originalStructures, ref orderedHeaders, ref isFirstItem);
                }
            }
            catch (JsonReaderException)
            {
                return Array.Empty<byte>();
            }

            if (flatList.Count == 0)
            {
                return Array.Empty<byte>();
            }

            return CreateGroupedFlattenedExcelWorkbook(flatList, originalStructures, orderedHeaders);
        }

        #region Enhanced Structure Analysis Methods

        /// <summary>
        /// Represents a record with both its flattened data and structure information
        /// </summary>
        private class StructuredRecord
        {
            public Dictionary<string, object> FlatData { get; set; }
            public StructureSignature Structure { get; set; }
            public JToken OriginalToken { get; set; }
        }

        /// <summary>
        /// Represents the structural signature of a JSON object
        /// </summary>
        private class StructureSignature
        {
            public HashSet<string> ObjectProperties { get; set; } = new HashSet<string>();
            public Dictionary<string, int> ArrayProperties { get; set; } = new Dictionary<string, int>();
            public HashSet<string> SimpleProperties { get; set; } = new HashSet<string>();
            public int TotalComplexity => ObjectProperties.Count + ArrayProperties.Count;

            public string GetSignatureKey()
            {
                var parts = new List<string>();

                // Add object properties
                if (ObjectProperties.Count > 0)
                {
                    parts.Add($"Objects[{string.Join(",", ObjectProperties.OrderBy(x => x))}]");
                }

                // Add array properties with their sizes
                if (ArrayProperties.Count > 0)
                {
                    var arrayParts = ArrayProperties.OrderBy(x => x.Key)
                        .Select(x => $"{x.Key}({x.Value})");
                    parts.Add($"Arrays[{string.Join(",", arrayParts)}]");
                }

                if (parts.Count == 0)
                {
                    return "SimpleRecord";
                }

                return string.Join("_", parts);
            }

            public string GetFriendlyName()
            {
                if (TotalComplexity == 0)
                    return "Simple Records";

                var nameParts = new List<string>();

                // Create meaningful names based on common patterns
                if (ObjectProperties.Contains("address") || ObjectProperties.Contains("Address"))
                    nameParts.Add("WithAddress");

                if (ObjectProperties.Contains("user") || ObjectProperties.Contains("User") ||
                    ObjectProperties.Contains("customer") || ObjectProperties.Contains("Customer"))
                    nameParts.Add("WithUser");

                if (ArrayProperties.ContainsKey("orders") || ArrayProperties.ContainsKey("Orders"))
                    nameParts.Add("WithOrders");

                if (ArrayProperties.ContainsKey("items") || ArrayProperties.ContainsKey("Items"))
                    nameParts.Add("WithItems");

                if (ObjectProperties.Contains("metadata") || ObjectProperties.Contains("Metadata"))
                    nameParts.Add("WithMetadata");

                if (nameParts.Count == 0)
                {
                    if (ObjectProperties.Count > 0 && ArrayProperties.Count > 0)
                        nameParts.Add("ComplexNested");
                    else if (ObjectProperties.Count > 0)
                        nameParts.Add("WithObjects");
                    else if (ArrayProperties.Count > 0)
                        nameParts.Add("WithArrays");
                }

                return nameParts.Count > 0 ? string.Join("_", nameParts) : "MixedStructure";
            }
        }

        /// <summary>
        /// Extracts records while maintaining detailed structure information
        /// </summary>
        private void ExtractRecordsWithStructure(JToken token, List<StructuredRecord> results)
        {
            if (token is JArray array)
            {
                foreach (var item in array)
                {
                    ExtractRecordsWithStructure(item, results);
                }
            }
            else if (token is JObject obj)
            {
                // Check if this object contains array properties that should be treated as separate records
                var arrayContainers = new[] { "Documents", "Items", "Data", "Records", "Results", "Entities" };

                foreach (var container in arrayContainers)
                {
                    if (obj[container] is JArray containerArray)
                    {
                        ExtractRecordsWithStructure(containerArray, results);
                        return;
                    }
                }

                // This is a leaf record - analyze its structure
                var signature = AnalyzeStructure(obj);
                var flatData = new Dictionary<string, object>();
                FlattenJsonToDictionary(obj, "", flatData);

                results.Add(new StructuredRecord
                {
                    FlatData = flatData,
                    Structure = signature,
                    OriginalToken = obj
                });
            }
        }

        /// <summary>
        /// Analyzes the structure of a JSON object to create a signature
        /// </summary>
        private StructureSignature AnalyzeStructure(JToken token)
        {
            var signature = new StructureSignature();

            if (token is JObject obj)
            {
                foreach (var prop in obj.Properties())
                {
                    switch (prop.Value.Type)
                    {
                        case JTokenType.Object:
                            signature.ObjectProperties.Add(prop.Name);
                            break;
                        case JTokenType.Array:
                            var array = prop.Value as JArray;
                            signature.ArrayProperties[prop.Name] = array?.Count ?? 0;
                            break;
                        default:
                            signature.SimpleProperties.Add(prop.Name);
                            break;
                    }
                }
            }

            return signature;
        }

        /// <summary>
        /// Groups records by their structural patterns
        /// </summary>
        private Dictionary<string, List<StructuredRecord>> GroupRecordsByStructure(List<StructuredRecord> records)
        {
            var groups = new Dictionary<string, List<StructuredRecord>>();

            foreach (var record in records)
            {
                var groupKey = record.Structure.GetFriendlyName();

                if (!groups.ContainsKey(groupKey))
                {
                    groups[groupKey] = new List<StructuredRecord>();
                }

                groups[groupKey].Add(record);
            }

            // If we only have one group, try to sub-group by specific properties
            if (groups.Count == 1 && records.Count > 10)
            {
                var singleGroup = groups.First();
                var subGroups = TryCreateSubGroups(singleGroup.Value);
                if (subGroups.Count > 1)
                {
                    groups.Clear();
                    foreach (var subGroup in subGroups)
                    {
                        groups[subGroup.Key] = subGroup.Value;
                    }
                }
            }

            return groups;
        }

        /// <summary>
        /// Attempts to create meaningful sub-groups from records with similar structures
        /// </summary>
        private Dictionary<string, List<StructuredRecord>> TryCreateSubGroups(List<StructuredRecord> records)
        {
            var subGroups = new Dictionary<string, List<StructuredRecord>>();

            // Try grouping by specific field patterns
            var groupingStrategies = new[]
            {
                // Group by document type or category
                new Func<Dictionary<string, object>, string>(data =>
                    GetValueAsString(data, "type", "documentType", "category", "kind")),
                
                // Group by status or state
                new Func<Dictionary<string, object>, string>(data =>
                    GetValueAsString(data, "status", "state", "condition")),
                
                // Group by presence of specific complex properties
                new Func<Dictionary<string, object>, string>(data =>
                {
                    if (data.Keys.Any(k => k.Contains("order") || k.Contains("Order")))
                        return "WithOrders";
                    if (data.Keys.Any(k => k.Contains("address") || k.Contains("Address")))
                        return "WithAddress";
                    if (data.Keys.Any(k => k.Contains("user") || k.Contains("User")))
                        return "WithUser";
                    return "Standard";
                })
            };

            foreach (var strategy in groupingStrategies)
            {
                var testGroups = new Dictionary<string, List<StructuredRecord>>();

                foreach (var record in records)
                {
                    var groupKey = strategy(record.FlatData) ?? "Unknown";

                    if (!testGroups.ContainsKey(groupKey))
                        testGroups[groupKey] = new List<StructuredRecord>();

                    testGroups[groupKey].Add(record);
                }

                // If this strategy creates meaningful groups (more than 1 group, no group too dominant)
                if (testGroups.Count > 1 && testGroups.Values.Max(g => g.Count) < records.Count * 0.8)
                {
                    return testGroups;
                }
            }

            // No meaningful sub-grouping found
            return new Dictionary<string, List<StructuredRecord>> { { "All Records", records } };
        }

        /// <summary>
        /// Gets a string value from data dictionary using multiple possible keys
        /// </summary>
        private string GetValueAsString(Dictionary<string, object> data, params string[] possibleKeys)
        {
            foreach (var key in possibleKeys)
            {
                if (data.TryGetValue(key, out var value) && value != null)
                {
                    return value.ToString();
                }

                // Try case-insensitive match
                var foundKey = data.Keys.FirstOrDefault(k =>
                    string.Equals(k, key, StringComparison.OrdinalIgnoreCase));

                if (foundKey != null && data[foundKey] != null)
                {
                    return data[foundKey].ToString();
                }
            }

            return null;
        }

        /// <summary>
        /// Flattens JSON to dictionary format (non-recursive for performance)
        /// </summary>
        private void FlattenJsonToDictionary(JToken token, string prefix, Dictionary<string, object> result)
        {
            if (token is JObject obj)
            {
                foreach (var prop in obj.Properties())
                {
                    var key = string.IsNullOrEmpty(prefix) ? prop.Name : $"{prefix}.{prop.Name}";

                    if (prop.Value is JObject || prop.Value is JArray)
                    {
                        FlattenJsonToDictionary(prop.Value, key, result);
                    }
                    else
                    {
                        result[key] = ((JValue)prop.Value)?.Value;
                    }
                }
            }
            else if (token is JArray array)
            {
                for (int i = 0; i < array.Count; i++)
                {
                    var key = $"{prefix}[{i}]";

                    if (array[i] is JObject || array[i] is JArray)
                    {
                        FlattenJsonToDictionary(array[i], key, result);
                    }
                    else
                    {
                        result[key] = ((JValue)array[i])?.Value;
                    }
                }
            }
        }

        /// <summary>
        /// Creates an enhanced summary sheet with detailed structure information
        /// </summary>
        private void CreateEnhancedSummarySheet(XLWorkbook workbook, Dictionary<string, List<StructuredRecord>> groups)
        {
            var summarySheet = workbook.Worksheets.Add("Summary");

            // Create title
            var titleCell = summarySheet.Cell(1, 1);
            titleCell.Value = "Export Summary - Grouped by Data Structure";
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Font.FontSize = 16;
            titleCell.Style.Fill.BackgroundColor = XLColor.LightBlue;
            titleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            summarySheet.Range(1, 1, 1, 5).Merge();

            // Create column headers
            var headers = new[] { "Structure Type", "Record Count", "Percentage", "Sample Properties", "Sheet Name" };
            for (int col = 0; col < headers.Length; col++)
            {
                summarySheet.Cell(3, col + 1).Value = headers[col];
            }

            var headerRange = summarySheet.Range(3, 1, 3, 5);
            StyleHeaderRange(headerRange);

            int row = 4;
            int totalRecords = groups.Values.Sum(g => g.Count);

            // Add group statistics with enhanced information
            foreach (var group in groups.OrderByDescending(g => g.Value.Count))
            {
                var count = group.Value.Count;
                var percentage = totalRecords > 0 ? (count * 100.0 / totalRecords) : 0;
                var sampleRecord = group.Value.FirstOrDefault();
                var sampleProps = sampleRecord?.FlatData.Keys.Take(3).ToArray() ?? new string[0];
                var propsText = sampleProps.Length > 0 ? string.Join(", ", sampleProps) : "N/A";

                summarySheet.Cell(row, 1).Value = group.Key;
                summarySheet.Cell(row, 2).Value = count;
                summarySheet.Cell(row, 3).Value = $"{percentage:F1}%";
                summarySheet.Cell(row, 4).Value = propsText;
                summarySheet.Cell(row, 5).Value = SanitizeSheetName(group.Key);
                row++;
            }

            // Style data rows
            if (row > 4)
            {
                var dataRange = summarySheet.Range(4, 1, row - 1, 5);
                StyleDataRange(dataRange);
            }

            // Add total row
            summarySheet.Cell(row + 1, 1).Value = "Total Records:";
            summarySheet.Cell(row + 1, 2).Value = totalRecords;
            summarySheet.Cell(row + 1, 1).Style.Font.Bold = true;
            summarySheet.Cell(row + 1, 2).Style.Font.Bold = true;

            // Add generation timestamp
            summarySheet.Cell(row + 3, 1).Value = $"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
            summarySheet.Cell(row + 3, 1).Style.Font.Italic = true;

            summarySheet.Columns().AdjustToContents();
        }

        /// <summary>
        /// Creates an error sheet when processing fails
        /// </summary>
        private byte[] CreateErrorSheet(string errorMessage)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Error");

            worksheet.Cell(1, 1).Value = "Export Error";
            worksheet.Cell(1, 1).Style.Font.Bold = true;
            worksheet.Cell(1, 1).Style.Font.FontSize = 14;
            worksheet.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.LightPink;

            worksheet.Cell(3, 1).Value = "Error Details:";
            worksheet.Cell(3, 1).Style.Font.Bold = true;

            worksheet.Cell(4, 1).Value = errorMessage;
            worksheet.Cell(4, 1).Style.Alignment.WrapText = true;

            worksheet.Cell(6, 1).Value = $"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
            worksheet.Cell(6, 1).Style.Font.Italic = true;

            worksheet.Columns().AdjustToContents();

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return stream.ToArray();
        }

        #endregion

        #region Existing Methods (Updated)

        /// <summary>
        /// Gets all unique headers from a list of data dictionaries
        /// </summary>
        private string[] GetAllUniqueHeaders(List<Dictionary<string, object>> dataList)
        {
            var allHeaders = new HashSet<string>();

            foreach (var item in dataList)
            {
                foreach (var key in item.Keys)
                {
                    allHeaders.Add(key);
                }
            }

            return allHeaders.OrderBy(h => h).ToArray();
        }

        /// <summary>
        /// Groups records by their nested content structure
        /// </summary>
        private Dictionary<string, List<Dictionary<string, object>>> GroupByNestedStructure(List<Dictionary<string, object>> dataList)
        {
            var groups = new Dictionary<string, List<Dictionary<string, object>>>();

            foreach (var item in dataList)
            {
                var structureKey = GetNestedStructureKey(item);

                if (!groups.ContainsKey(structureKey))
                {
                    groups[structureKey] = new List<Dictionary<string, object>>();
                }

                groups[structureKey].Add(item);
            }

            return groups;
        }

        /// <summary>
        /// Generates a key representing the nested structure of a record
        /// Enhanced to handle more complex Cosmos DB structures
        /// </summary>
        private string GetNestedStructureKey(Dictionary<string, object> item)
        {
            var nestedProperties = new List<string>();
            var hasArrays = false;
            var hasObjects = false;

            foreach (var kvp in item)
            {
                if (kvp.Value is Newtonsoft.Json.Linq.JObject)
                {
                    hasObjects = true;
                    nestedProperties.Add($"Obj_{kvp.Key}");
                }
                else if (kvp.Value is Newtonsoft.Json.Linq.JArray jArray)
                {
                    hasArrays = true;
                    nestedProperties.Add($"Array_{kvp.Key}({jArray.Count})");
                }
                else if (kvp.Value != null && kvp.Value.GetType().IsClass && kvp.Value.GetType() != typeof(string))
                {
                    // Handle other complex types that might come from Cosmos DB
                    hasObjects = true;
                    nestedProperties.Add($"Complex_{kvp.Key}");
                }
                else if (kvp.Value is Dictionary<string, object>)
                {
                    hasObjects = true;
                    nestedProperties.Add($"Dict_{kvp.Key}");
                }
            }

            if (nestedProperties.Count == 0)
            {
                return "SimpleRecords";
            }

            // Create more readable group names based on common Cosmos DB patterns
            var keyLower = string.Join("_", nestedProperties).ToLower();

            if (keyLower.Contains("orders") && keyLower.Contains("address"))
                return "WithOrdersAndAddress";
            else if (keyLower.Contains("orders") && keyLower.Contains("customer"))
                return "WithOrdersAndCustomer";
            else if (keyLower.Contains("user") && keyLower.Contains("profile"))
                return "WithUserProfiles";
            else if (keyLower.Contains("orders"))
                return "WithOrders";
            else if (keyLower.Contains("membership") || keyLower.Contains("subscription"))
                return "WithMembership";
            else if (keyLower.Contains("address") || keyLower.Contains("location"))
                return "WithAddress";
            else if (keyLower.Contains("metadata") || keyLower.Contains("properties"))
                return "WithMetadata";
            else if (keyLower.Contains("items") && hasArrays)
                return "WithItemArrays";
            else if (hasArrays && hasObjects)
                return "ComplexNested";
            else if (hasArrays)
                return "WithArrays";
            else if (hasObjects)
                return "WithObjects";

            var key = string.Join("_", nestedProperties.OrderBy(p => p));
            return key.Length > 31 ? key.Substring(0, 31) : key;
        }

        /// <summary>
        /// Processes JSON array while tracking original structure information
        /// </summary>
        private void ProcessJsonArrayWithStructure(JArray array, List<Dictionary<string, object>> flatList,
            List<Dictionary<string, object>> originalStructures, ref List<string> orderedHeaders, ref bool isFirstItem)
        {
            foreach (var item in array)
            {
                flatList.Add(FlattenJson(item, "", isFirstItem, ref orderedHeaders));
                originalStructures.Add(ExtractStructureInfo(item));
                isFirstItem = false;
            }
        }

        /// <summary>
        /// Extracts structure information from a JSON token
        /// Enhanced for better Cosmos DB structure detection
        /// </summary>
        private Dictionary<string, object> ExtractStructureInfo(JToken token)
        {
            var structure = new Dictionary<string, object>();

            if (token is JObject obj)
            {
                foreach (var prop in obj.Properties())
                {
                    switch (prop.Value.Type)
                    {
                        case JTokenType.Object:
                            structure[$"has_{prop.Name}"] = "object";
                            // Also capture nested object property names for better grouping
                            if (prop.Value is JObject nestedObj)
                            {
                                var nestedProps = nestedObj.Properties().Select(p => p.Name).Take(3);
                                structure[$"nested_{prop.Name}_props"] = string.Join(",", nestedProps);
                            }
                            break;
                        case JTokenType.Array:
                            var arr = prop.Value as JArray;
                            structure[$"has_{prop.Name}"] = $"array[{arr?.Count ?? 0}]";
                            // Capture array element type for better classification
                            if (arr?.Count > 0)
                            {
                                var firstElement = arr[0];
                                structure[$"array_{prop.Name}_type"] = firstElement.Type.ToString().ToLower();
                            }
                            break;
                        default:
                            structure[$"simple_{prop.Name}"] = prop.Value.Type.ToString().ToLower();
                            break;
                    }
                }
            }

            return structure;
        }

        /// <summary>
        /// Creates grouped Excel workbook from flattened data based on nested structure
        /// </summary>
        private byte[] CreateGroupedFlattenedExcelWorkbook(List<Dictionary<string, object>> flatList,
            List<Dictionary<string, object>> originalStructures, List<string> orderedHeaders)
        {
            using var workbook = new XLWorkbook();

            // Build complete header list
            var allHeaders = new HashSet<string>(orderedHeaders);
            foreach (var item in flatList)
            {
                foreach (var key in item.Keys)
                {
                    allHeaders.Add(key);
                }
            }

            var finalHeaders = orderedHeaders.Concat(allHeaders.Except(orderedHeaders).OrderBy(h => h)).ToArray();

            // Group by nested structure patterns
            var structureGroups = new Dictionary<string, List<int>>();

            for (int i = 0; i < originalStructures.Count; i++)
            {
                var structureKey = GetStructureKeyFromInfo(originalStructures[i]);

                if (!structureGroups.ContainsKey(structureKey))
                {
                    structureGroups[structureKey] = new List<int>();
                }

                structureGroups[structureKey].Add(i);
            }

            // Create sheets for each structure group
            foreach (var group in structureGroups.OrderBy(g => g.Key))
            {
                var groupData = group.Value.Select(index => flatList[index]).ToList();
                var sanitizedName = SanitizeSheetName(group.Key);
                CreateWorksheet(workbook, sanitizedName, groupData, finalHeaders);
            }

            // Create summary sheet
            CreateStructureGroupSummarySheet(workbook, structureGroups, "NestedStructure");

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return stream.ToArray();
        }

        /// <summary>
        /// Gets structure key from extracted structure info
        /// Enhanced with better pattern recognition for Cosmos DB data
        /// </summary>
        private string GetStructureKeyFromInfo(Dictionary<string, object> structureInfo)
        {
            if (structureInfo.Count == 0)
                return "SimpleRecord";

            var parts = structureInfo.Keys.OrderBy(k => k).ToList();
            var values = structureInfo.Values.Select(v => v.ToString().ToLower()).ToList();
            var allText = string.Join(" ", parts.Concat(values)).ToLower();

            // Enhanced pattern recognition for common Cosmos DB structures
            if (allText.Contains("orders") && allText.Contains("address"))
                return "WithOrdersAndAddress";
            else if (allText.Contains("orders") && allText.Contains("customer"))
                return "WithOrdersAndCustomer";
            else if (allText.Contains("user") && allText.Contains("profile"))
                return "WithUserProfiles";
            else if (allText.Contains("product") && allText.Contains("category"))
                return "WithProductCategories";
            else if (allText.Contains("orders") || allText.Contains("transactions"))
                return "WithOrders";
            else if (allText.Contains("membership") || allText.Contains("subscription"))
                return "WithMembership";
            else if (allText.Contains("address") || allText.Contains("location") || allText.Contains("coordinates"))
                return "WithAddress";
            else if (allText.Contains("metadata") || allText.Contains("properties") || allText.Contains("attributes"))
                return "WithMetadata";
            else if (allText.Contains("items") && allText.Contains("array"))
                return "WithItemArrays";
            else if (allText.Contains("tags") && allText.Contains("array"))
                return "WithTagArrays";
            else if (allText.Contains("array") && allText.Contains("object"))
                return "ComplexNested";
            else if (allText.Contains("array"))
                return "WithArrays";
            else if (allText.Contains("object"))
                return "WithObjects";

            return "MixedStructure";
        }

        /// <summary>
        /// Helper method to process JSON arrays efficiently
        /// </summary>
        private void ProcessJsonArray(JArray array, List<Dictionary<string, object>> flatList,
            ref List<string> orderedHeaders, ref bool isFirstItem)
        {
            foreach (var item in array)
            {
                flatList.Add(FlattenJson(item, "", isFirstItem, ref orderedHeaders));
                isFirstItem = false;
            }
        }

        /// <summary>
        /// Creates Excel workbook from flattened data with optimized performance
        /// </summary>
        private byte[] CreateFlattenedExcelWorkbook(List<Dictionary<string, object>> flatList, List<string> orderedHeaders)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Data");

            // Combine ordered headers with any additional headers found in other records
            var allHeaders = new HashSet<string>(orderedHeaders);
            foreach (var item in flatList)
            {
                foreach (var key in item.Keys)
                {
                    allHeaders.Add(key);
                }
            }

            var finalHeaders = orderedHeaders.Concat(allHeaders.Except(orderedHeaders).OrderBy(h => h)).ToArray();

            // Bulk create headers
            for (int col = 0; col < finalHeaders.Length; col++)
            {
                worksheet.Cell(1, col + 1).Value = finalHeaders[col];
            }

            var headerRange = worksheet.Range(1, 1, 1, finalHeaders.Length);
            StyleHeaderRange(headerRange);
            headerRange.SetAutoFilter();
            worksheet.SheetView.FreezeRows(1);

            // Bulk insert data
            for (int row = 0; row < flatList.Count; row++)
            {
                var rowData = flatList[row];
                for (int col = 0; col < finalHeaders.Length; col++)
                {
                    if (rowData.TryGetValue(finalHeaders[col], out var value))
                    {
                        var cell = worksheet.Cell(row + 2, col + 1);
                        SetCellValueWithType(cell, value);
                    }
                }
            }

            // Style all data cells at once
            if (flatList.Count > 0)
            {
                var dataRange = worksheet.Range(2, 1, flatList.Count + 1, finalHeaders.Length);
                StyleDataRange(dataRange);
            }

            worksheet.Columns().AdjustToContents();

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return stream.ToArray();
        }

        /// <summary>
        /// Recursively flattens nested JSON objects and arrays into dot-notation keys
        /// Example: {user: {name: "John"}} becomes {"user.name": "John"}
        /// Enhanced for better Cosmos DB structure handling
        /// </summary>
        private Dictionary<string, object> FlattenJson(JToken token, string parentPath, bool isFirstItem, ref List<string> orderedHeaders)
        {
            var result = new Dictionary<string, object>();

            if (token is JObject jObj)
            {
                foreach (var prop in jObj.Properties())
                {
                    string path = string.IsNullOrEmpty(parentPath) ? prop.Name : $"{parentPath}.{prop.Name}";
                    var value = prop.Value;

                    switch (value.Type)
                    {
                        case JTokenType.Object:
                            // Recursively flatten nested objects
                            var nestedResult = FlattenJson(value, path, isFirstItem, ref orderedHeaders);
                            foreach (var nested in nestedResult)
                            {
                                result[nested.Key] = nested.Value;
                                if (isFirstItem && !orderedHeaders.Contains(nested.Key))
                                    orderedHeaders.Add(nested.Key);
                            }
                            break;

                        case JTokenType.Array:
                            // Handle arrays by creating indexed keys
                            var array = value as JArray;

                            // For small arrays, expand each element
                            if (array.Count <= 10)
                            {
                                for (int i = 0; i < array.Count; i++)
                                {
                                    if (array[i] is JObject itemObj)
                                    {
                                        var arrayResult = FlattenJson(itemObj, $"{path}[{i}]", isFirstItem, ref orderedHeaders);
                                        foreach (var nested in arrayResult)
                                        {
                                            result[nested.Key] = nested.Value;
                                            if (isFirstItem && !orderedHeaders.Contains(nested.Key))
                                                orderedHeaders.Add(nested.Key);
                                        }
                                    }
                                    else
                                    {
                                        string arrayKey = $"{path}[{i}]";
                                        result[arrayKey] = ((JValue)array[i]).Value;
                                        if (isFirstItem && !orderedHeaders.Contains(arrayKey))
                                            orderedHeaders.Add(arrayKey);
                                    }
                                }
                            }
                            else
                            {
                                // For large arrays, create summary information
                                result[$"{path}_Count"] = array.Count;
                                if (isFirstItem && !orderedHeaders.Contains($"{path}_Count"))
                                    orderedHeaders.Add($"{path}_Count");

                                // Sample first few elements
                                for (int i = 0; i < Math.Min(3, array.Count); i++)
                                {
                                    if (array[i] is JObject itemObj)
                                    {
                                        var arrayResult = FlattenJson(itemObj, $"{path}[{i}]", isFirstItem, ref orderedHeaders);
                                        foreach (var nested in arrayResult)
                                        {
                                            result[nested.Key] = nested.Value;
                                            if (isFirstItem && !orderedHeaders.Contains(nested.Key))
                                                orderedHeaders.Add(nested.Key);
                                        }
                                    }
                                    else
                                    {
                                        string arrayKey = $"{path}[{i}]";
                                        result[arrayKey] = ((JValue)array[i]).Value;
                                        if (isFirstItem && !orderedHeaders.Contains(arrayKey))
                                            orderedHeaders.Add(arrayKey);
                                    }
                                }
                            }
                            break;

                        case JTokenType.Null:
                            result[path] = null;
                            if (isFirstItem && !orderedHeaders.Contains(path))
                                orderedHeaders.Add(path);
                            break;

                        default:
                            // Simple values - just add them
                            result[path] = ((JValue)value).Value;
                            if (isFirstItem && !orderedHeaders.Contains(path))
                                orderedHeaders.Add(path);
                            break;
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Creates a worksheet with the provided data and styling
        /// </summary>
        public void CreateWorksheet(XLWorkbook workbook, string sheetName, List<Dictionary<string, object>> data, IEnumerable<string> headers)
        {
            var worksheet = workbook.Worksheets.Add(sheetName);
            var headerArray = headers.ToArray();

            if (data.Count == 0)
            {
                worksheet.Cell(1, 1).Value = "No data available";
                return;
            }

            // Create headers efficiently
            for (int col = 0; col < headerArray.Length; col++)
            {
                worksheet.Cell(1, col + 1).Value = headerArray[col];
            }

            var headerRange = worksheet.Range(1, 1, 1, headerArray.Length);
            StyleHeaderRange(headerRange);
            headerRange.SetAutoFilter();
            worksheet.SheetView.FreezeRows(1);

            // Insert data in bulk
            for (int row = 0; row < data.Count; row++)
            {
                var rowData = data[row];
                for (int col = 0; col < headerArray.Length; col++)
                {
                    var key = headerArray[col];
                    var cell = worksheet.Cell(row + 2, col + 1);

                    if (rowData.TryGetValue(key, out var value))
                    {
                        SetCellValueWithType(cell, value);
                    }
                    else
                    {
                        cell.Value = "";
                    }
                }
            }

            // Style all data cells at once for better performance
            if (data.Count > 0)
            {
                var dataRange = worksheet.Range(2, 1, data.Count + 1, headerArray.Length);
                StyleDataRange(dataRange);
            }

            worksheet.Columns().AdjustToContents();
        }

        /// <summary>
        /// Creates a summary sheet for nested structure groups
        /// </summary>
        private void CreateNestedStructureSummarySheet(XLWorkbook workbook, Dictionary<string, List<Dictionary<string, object>>> groups)
        {
            var summarySheet = workbook.Worksheets.Add("Summary");

            // Create title
            var titleCell = summarySheet.Cell(1, 1);
            titleCell.Value = "Export Summary - Grouped by Nested Structure";
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Font.FontSize = 14;
            titleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            summarySheet.Range(1, 1, 1, 3).Merge();

            // Create column headers
            var headers = new[] { "Structure Type", "Record Count", "Sheet Name" };
            for (int col = 0; col < headers.Length; col++)
            {
                summarySheet.Cell(3, col + 1).Value = headers[col];
            }

            var headerRange = summarySheet.Range(3, 1, 3, 3);
            StyleHeaderRange(headerRange);

            int row = 4;
            int totalRecords = 0;

            // Add group statistics
            foreach (var group in groups.OrderBy(g => g.Key))
            {
                summarySheet.Cell(row, 1).Value = group.Key;
                summarySheet.Cell(row, 2).Value = group.Value.Count;
                summarySheet.Cell(row, 3).Value = SanitizeSheetName(group.Key);
                totalRecords += group.Value.Count;
                row++;
            }

            // Style data rows
            if (row > 4)
            {
                var dataRange = summarySheet.Range(4, 1, row - 1, 3);
                StyleDataRange(dataRange);
            }

            // Add total row
            summarySheet.Cell(row + 1, 1).Value = "Total Records:";
            summarySheet.Cell(row + 1, 2).Value = totalRecords;
            summarySheet.Cell(row + 1, 1).Style.Font.Bold = true;
            summarySheet.Cell(row + 1, 2).Style.Font.Bold = true;

            summarySheet.Columns().AdjustToContents();
        }

        /// <summary>
        /// Creates summary sheet for structure-based groups
        /// </summary>
        private void CreateStructureGroupSummarySheet(XLWorkbook workbook, Dictionary<string, List<int>> groups, string groupByField)
        {
            var summarySheet = workbook.Worksheets.Add("Summary");

            var titleCell = summarySheet.Cell(1, 1);
            titleCell.Value = $"Export Summary - Grouped by {groupByField}";
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Font.FontSize = 14;
            titleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            summarySheet.Range(1, 1, 1, 3).Merge();

            summarySheet.Cell(3, 1).Value = "Structure Type";
            summarySheet.Cell(3, 2).Value = "Record Count";
            summarySheet.Cell(3, 3).Value = "Sheet Name";

            var headerRange = summarySheet.Range(3, 1, 3, 3);
            StyleHeaderRange(headerRange);

            int row = 4;
            int totalRecords = 0;

            foreach (var group in groups.OrderBy(g => g.Key))
            {
                summarySheet.Cell(row, 1).Value = group.Key;
                summarySheet.Cell(row, 2).Value = group.Value.Count;
                summarySheet.Cell(row, 3).Value = SanitizeSheetName(group.Key);
                totalRecords += group.Value.Count;
                row++;
            }

            if (row > 4)
            {
                var dataRange = summarySheet.Range(4, 1, row - 1, 3);
                StyleDataRange(dataRange);
            }

            summarySheet.Cell(row + 1, 1).Value = "Total Records:";
            summarySheet.Cell(row + 1, 2).Value = totalRecords;
            summarySheet.Cell(row + 1, 1).Style.Font.Bold = true;
            summarySheet.Cell(row + 1, 2).Style.Font.Bold = true;

            summarySheet.Columns().AdjustToContents();
        }

        /// <summary>
        /// Sanitizes sheet names to comply with Excel naming rules
        /// </summary>
        private string SanitizeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Unknown";

            // Replace invalid characters with underscore
            var invalidChars = new char[] { '/', '\\', '?', '*', '[', ']', ':' };
            var sb = new StringBuilder(name);

            foreach (var invalidChar in invalidChars)
            {
                sb.Replace(invalidChar, '_');
            }

            // Truncate if too long (Excel limit is 31 characters)
            var result = sb.ToString();
            if (result.Length > 31)
                result = result.Substring(0, 31);

            return result;
        }

        /// <summary>
        /// Applies consistent header styling to a range of cells
        /// </summary>
        private void StyleHeaderRange(IXLRange range)
        {
            range.Style.Font.Bold = true;
            range.Style.Fill.BackgroundColor = XLColor.LightGray;
            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            range.Style.Alignment.WrapText = true;
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        /// <summary>
        /// Applies consistent data styling to individual cells (kept for backward compatibility)
        /// </summary>
        private void StyleHeaderCell(IXLCell cell)
        {
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.LightGray;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Alignment.WrapText = true;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        /// <summary>
        /// Applies consistent data styling to a range of cells
        /// </summary>
        private void StyleDataRange(IXLRange range)
        {
            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            range.Style.Alignment.WrapText = true;
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        /// <summary>
        /// Applies consistent data styling to individual cells (kept for backward compatibility)
        /// </summary>
        private void StyleDataCell(IXLCell cell)
        {
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Alignment.WrapText = true;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        /// <summary>
        /// Intelligently sets cell values with proper type detection and formatting
        /// Uses caching to avoid repeated type parsing for better performance
        /// </summary>
        private void SetCellValueWithType(IXLCell cell, object value)
        {
            if (value == null)
            {
                cell.Value = "";
                return;
            }

            // Handle direct type matches first (fastest path)
            switch (value)
            {
                case bool b:
                    cell.Value = b;
                    return;
                case int i:
                    cell.Value = i;
                    return;
                case long l:
                    cell.Value = l;
                    return;
                case double d:
                    cell.Value = d;
                    return;
                case decimal m:
                    cell.Value = m;
                    return;
                case DateTime dt:
                    cell.Value = dt;
                    cell.Style.DateFormat.Format = "yyyy-mm-dd hh:mm:ss";
                    return;
                case Guid g:
                    cell.Value = g.ToString();
                    return;
            }

            // For string values, try to parse to appropriate types
            var str = value.ToString();

            // Use cached type information if available
            if (_typeCache.TryGetValue(str, out var cachedType))
            {
                SetCellValueFromCachedType(cell, str, cachedType);
                return;
            }

            // Try parsing in order of likelihood for performance
            if (bool.TryParse(str, out var boolParsed))
            {
                cell.Value = boolParsed;
                _typeCache[str] = typeof(bool);
            }
            else if (int.TryParse(str, out var intParsed))
            {
                cell.Value = intParsed;
                _typeCache[str] = typeof(int);
            }
            else if (long.TryParse(str, out var longParsed))
            {
                cell.Value = longParsed;
                _typeCache[str] = typeof(long);
            }
            else if (decimal.TryParse(str, out var decimalParsed))
            {
                cell.Value = decimalParsed;
                _typeCache[str] = typeof(decimal);
            }
            else if (DateTime.TryParse(str, out var dtParsed))
            {
                cell.Value = dtParsed;
                cell.Style.DateFormat.Format = "yyyy-mm-dd hh:mm:ss";
                _typeCache[str] = typeof(DateTime);
            }
            else if (Guid.TryParse(str, out var guidParsed))
            {
                cell.Value = str; // Keep as string for better readability
                _typeCache[str] = typeof(Guid);
            }
            else
            {
                cell.Value = str;
                _typeCache[str] = typeof(string);
            }
        }

        /// <summary>
        /// Sets cell value using cached type information for better performance
        /// </summary>
        private void SetCellValueFromCachedType(IXLCell cell, string str, Type type)
        {
            if (type == typeof(bool))
            {
                cell.Value = bool.Parse(str);
            }
            else if (type == typeof(int))
            {
                cell.Value = int.Parse(str);
            }
            else if (type == typeof(long))
            {
                cell.Value = long.Parse(str);
            }
            else if (type == typeof(decimal))
            {
                cell.Value = decimal.Parse(str);
            }
            else if (type == typeof(DateTime))
            {
                cell.Value = DateTime.Parse(str);
                cell.Style.DateFormat.Format = "yyyy-mm-dd hh:mm:ss";
            }
            else if (type == typeof(Guid))
            {
                cell.Value = str; // Keep GUIDs as strings for readability
            }
            else
            {
                cell.Value = str;
            }
        }

        #endregion
    }




    /// <summary>
    /// High-performance Excel to JSON converter with intelligent type detection
    /// and support for complex nested structures
    /// </summary>
    public class ExcelToJsonConverter
    {
        private readonly ExcelImportOptions _options;
        private readonly TypeDetector _typeDetector;
        private readonly StructureAnalyzer _structureAnalyzer;

        public ExcelToJsonConverter(ExcelImportOptions options = null)
        {
            _options = options ?? new ExcelImportOptions();
            _typeDetector = new TypeDetector(_options);
            _structureAnalyzer = new StructureAnalyzer();
        }

        #region Public Methods

        /// <summary>
        /// Converts Excel file bytes to list of data with automatic structure detection
        /// </summary>
        public async Task<List<Dictionary<string, object>>> ConvertToDataAsync(byte[] excelBytes, ConversionMode mode = ConversionMode.Auto)
        {
            return await Task.Run(() => ConvertToData(excelBytes, mode));
        }

        /// <summary>
        /// Synchronous conversion of Excel file bytes to list of data
        /// </summary>
        public List<Dictionary<string, object>> ConvertToData(byte[] excelBytes, ConversionMode mode = ConversionMode.Auto)
        {
            try
            {
                using var stream = new MemoryStream(excelBytes);
                using var workbook = new XLWorkbook(stream);

                if (workbook.Worksheets.Count == 0)
                {
                    throw new InvalidOperationException("Excel file contains no worksheets");
                }

                // Determine conversion mode if auto
                if (mode == ConversionMode.Auto)
                {
                    mode = DetermineOptimalMode(workbook);
                }

                // Process based on mode and return list of data
                return mode switch
                {
                    ConversionMode.Simple => ConvertSimple(workbook),
                    ConversionMode.Nested => ConvertNested(workbook),
                    ConversionMode.MultiSheet => ConvertMultiSheetToList(workbook),
                    ConversionMode.Grouped => ConvertGroupedToList(workbook),
                    _ => ConvertSimple(workbook)
                };
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Conversion failed: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Converts Excel file from stream to list of data
        /// </summary>
        public async Task<List<Dictionary<string, object>>> ConvertFromStreamAsync(Stream stream, ConversionMode mode = ConversionMode.Auto)
        {
            using var memoryStream = new MemoryStream();
            await stream.CopyToAsync(memoryStream);
            return await ConvertToDataAsync(memoryStream.ToArray(), mode);
        }

        /// <summary>
        /// Converts Excel file from file path to list of data
        /// </summary>
        public async Task<List<Dictionary<string, object>>> ConvertFromFileAsync(string filePath, ConversionMode mode = ConversionMode.Auto)
        {
            var bytes = await File.ReadAllBytesAsync(filePath);
            return await ConvertToDataAsync(bytes, mode);
        }

        /// <summary>
        /// Legacy method - returns ConversionResult for backward compatibility
        /// </summary>
        public ConversionResult ConvertToJson(byte[] excelBytes, ConversionMode mode = ConversionMode.Auto)
        {
            var result = new ConversionResult();

            try
            {
                var data = ConvertToData(excelBytes, mode);

                result.Success = true;
                result.Data = data;
                result.JsonString = JsonConvert.SerializeObject(data, _options.JsonFormatting);
                result.RecordCount = data.Count;
                result.ConversionMode = mode;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Conversion failed: {ex.Message}";
                result.Exception = ex;
            }

            return result;
        }

        #endregion

        #region Conversion Modes

        private List<Dictionary<string, object>> ConvertSimple(XLWorkbook workbook)
        {
            var worksheet = workbook.Worksheets.First();
            return ExtractSheetData(worksheet);
        }

        private List<Dictionary<string, object>> ConvertNested(XLWorkbook workbook)
        {
            var worksheet = workbook.Worksheets.First();
            var flatData = ExtractSheetData(worksheet);
            return ReconstructNestedStructure(flatData);
        }

        private List<Dictionary<string, object>> ConvertMultiSheetToList(XLWorkbook workbook)
        {
            var allData = new List<Dictionary<string, object>>();

            foreach (var worksheet in workbook.Worksheets)
            {
                if (!ShouldSkipSheet(worksheet))
                {
                    var sheetData = ExtractSheetData(worksheet);

                    // Add sheet name as metadata if configured
                    if (_options.IncludeSheetMetadata)
                    {
                        foreach (var record in sheetData)
                        {
                            record["_sheetName"] = worksheet.Name;
                        }
                    }

                    allData.AddRange(sheetData);
                }
            }

            return allData;
        }

        private List<Dictionary<string, object>> ConvertGroupedToList(XLWorkbook workbook)
        {
            var allData = new List<Dictionary<string, object>>();

            // Extract data from all sheets
            foreach (var worksheet in workbook.Worksheets)
            {
                if (!ShouldSkipSheet(worksheet))
                {
                    var sheetData = ExtractSheetData(worksheet);

                    // Add sheet name as metadata if configured
                    if (_options.IncludeSheetMetadata)
                    {
                        foreach (var record in sheetData)
                        {
                            record["_sheetName"] = worksheet.Name;
                        }
                    }

                    allData.AddRange(sheetData);
                }
            }

            // For grouped mode, you might want to add group information
            var groupedData = _structureAnalyzer.GroupByStructure(allData);
            var flattenedData = new List<Dictionary<string, object>>();

            foreach (var group in groupedData)
            {
                foreach (var record in group.Value)
                {
                    if (_options.IncludeSheetMetadata)
                    {
                        record["_groupName"] = group.Key;
                    }
                    flattenedData.Add(record);
                }
            }

            return flattenedData;
        }

        #endregion

        #region Data Extraction

        private List<Dictionary<string, object>> ExtractSheetData(IXLWorksheet worksheet)
        {
            var data = new List<Dictionary<string, object>>();
            var usedRange = worksheet.RangeUsed();

            if (usedRange == null)
                return data;

            // Extract headers
            var headers = ExtractHeaders(worksheet, usedRange);
            if (headers.Count == 0)
                return data;

            // Extract data rows
            var firstDataRow = _options.HeaderRow + 1;
            var lastRow = usedRange.LastRow().RowNumber();

            for (int row = firstDataRow; row <= lastRow; row++)
            {
                var rowData = new Dictionary<string, object>();
                bool hasData = false;

                for (int col = 0; col < headers.Count; col++)
                {
                    var cell = worksheet.Cell(row, col + 1);
                    var value = ExtractCellValue(cell);

                    if (value != null && !string.IsNullOrWhiteSpace(value.ToString()))
                    {
                        hasData = true;
                    }

                    rowData[headers[col]] = value;
                }

                // Only add rows that contain data
                if (hasData || !_options.SkipEmptyRows)
                {
                    data.Add(rowData);
                }
            }

            return data;
        }

        private List<string> ExtractHeaders(IXLWorksheet worksheet, IXLRange usedRange)
        {
            var headers = new List<string>();
            var headerRow = worksheet.Row(_options.HeaderRow);
            var lastColumn = usedRange.LastColumn().ColumnNumber();

            for (int col = 1; col <= lastColumn; col++)
            {
                var cell = headerRow.Cell(col);
                var headerValue = cell.GetValue<string>()?.Trim();

                if (string.IsNullOrWhiteSpace(headerValue))
                {
                    if (_options.GenerateMissingHeaders)
                    {
                        headerValue = $"Column{col}";
                    }
                    else
                    {
                        continue;
                    }
                }

                // Ensure unique headers
                var finalHeader = headerValue;
                int suffix = 1;
                while (headers.Contains(finalHeader))
                {
                    finalHeader = $"{headerValue}_{suffix++}";
                }

                headers.Add(finalHeader);
            }

            return headers;
        }

        private object ExtractCellValue(IXLCell cell)
        {
            if (cell.IsEmpty())
                return null;

            // Check for formula results first
            if (cell.HasFormula && _options.EvaluateFormulas)
            {
                try
                {
                    return cell.CachedValue;
                }
                catch
                {
                    // Formula evaluation failed, try getting raw value
                }
            }

            // Use type detector for intelligent type conversion
            return _typeDetector.DetectAndConvert(cell);
        }

        #endregion

        #region Nested Structure Reconstruction

        private List<Dictionary<string, object>> ReconstructNestedStructure(List<Dictionary<string, object>> flatData)
        {
            var result = new List<Dictionary<string, object>>();

            foreach (var flatRecord in flatData)
            {
                var nestedRecord = new Dictionary<string, object>();

                foreach (var kvp in flatRecord)
                {
                    SetNestedValue(nestedRecord, kvp.Key, kvp.Value);
                }

                result.Add(nestedRecord);
            }

            return result;
        }

        private void SetNestedValue(Dictionary<string, object> target, string path, object value)
        {
            var parts = ParsePath(path);
            var current = target;

            for (int i = 0; i < parts.Count - 1; i++)
            {
                var part = parts[i];

                if (part.IsArray)
                {
                    // Handle array notation
                    if (!current.ContainsKey(part.Name))
                    {
                        current[part.Name] = new List<Dictionary<string, object>>();
                    }

                    var list = current[part.Name] as List<Dictionary<string, object>>;

                    // Ensure list has enough elements
                    while (list.Count <= part.Index)
                    {
                        list.Add(new Dictionary<string, object>());
                    }

                    current = list[part.Index];
                }
                else
                {
                    // Handle object notation
                    if (!current.ContainsKey(part.Name))
                    {
                        current[part.Name] = new Dictionary<string, object>();
                    }

                    current = current[part.Name] as Dictionary<string, object>;
                }
            }

            // Set the final value
            var lastPart = parts.Last();
            if (lastPart.IsArray)
            {
                if (!current.ContainsKey(lastPart.Name))
                {
                    current[lastPart.Name] = new List<object>();
                }

                var list = current[lastPart.Name] as List<object>;
                while (list.Count <= lastPart.Index)
                {
                    list.Add(null);
                }

                list[lastPart.Index] = value;
            }
            else
            {
                current[lastPart.Name] = value;
            }
        }

        private List<PathPart> ParsePath(string path)
        {
            var parts = new List<PathPart>();
            var regex = new Regex(@"([^\.\[]+)(\[(\d+)\])?");
            var segments = path.Split('.');

            foreach (var segment in segments)
            {
                var matches = regex.Matches(segment);
                foreach (Match match in matches)
                {
                    var part = new PathPart
                    {
                        Name = match.Groups[1].Value
                    };

                    if (match.Groups[3].Success)
                    {
                        part.IsArray = true;
                        part.Index = int.Parse(match.Groups[3].Value);
                    }

                    parts.Add(part);
                }
            }

            return parts;
        }

        #endregion

        #region Helper Methods

        private ConversionMode DetermineOptimalMode(XLWorkbook workbook)
        {
            // Multiple sheets suggest multi-sheet mode
            if (workbook.Worksheets.Count > 1)
            {
                // Check if sheets have similar structure
                var structures = new List<HashSet<string>>();

                foreach (var sheet in workbook.Worksheets)
                {
                    if (!ShouldSkipSheet(sheet))
                    {
                        var headers = ExtractHeaders(sheet, sheet.RangeUsed());
                        structures.Add(new HashSet<string>(headers));
                    }
                }

                // If structures are similar, consider grouped mode
                if (structures.Count > 1 && AreSimilarStructures(structures))
                {
                    return ConversionMode.Grouped;
                }

                return ConversionMode.MultiSheet;
            }

            // Single sheet - check for nested structure indicators
            var worksheet = workbook.Worksheets.First();
            var usedRange = worksheet.RangeUsed();

            if (usedRange != null)
            {
                var headers = ExtractHeaders(worksheet, usedRange);

                // Check for dot notation or array notation in headers
                if (headers.Any(h => h.Contains(".") || h.Contains("[") || h.Contains("]")))
                {
                    return ConversionMode.Nested;
                }
            }

            return ConversionMode.Simple;
        }

        private bool AreSimilarStructures(List<HashSet<string>> structures)
        {
            if (structures.Count < 2)
                return false;

            var first = structures[0];

            foreach (var structure in structures.Skip(1))
            {
                var intersection = first.Intersect(structure).Count();
                var union = first.Union(structure).Count();

                // Consider similar if 70% overlap
                if (intersection / (double)union < 0.7)
                    return false;
            }

            return true;
        }

        private bool ShouldSkipSheet(IXLWorksheet worksheet)
        {
            // Skip summary sheets or empty sheets
            var name = worksheet.Name.ToLower();

            if (name == "summary" || name == "index" || name == "toc")
                return true;

            var usedRange = worksheet.RangeUsed();
            return usedRange == null || usedRange.RowCount() == 0;
        }

        #endregion

        #region Nested Classes

        private class PathPart
        {
            public string Name { get; set; }
            public bool IsArray { get; set; }
            public int Index { get; set; }
        }

        #endregion
    }

    /// <summary>
    /// Intelligent type detection and conversion for Excel cells
    /// </summary>
    public class TypeDetector
    {
        private readonly ExcelImportOptions _options;
        private readonly Dictionary<string, Type> _typeCache;
        private readonly Dictionary<string, string> _dateFormats;

        public TypeDetector(ExcelImportOptions options)
        {
            _options = options;
            _typeCache = new Dictionary<string, Type>();
            _dateFormats = new Dictionary<string, string>
        {
            { "yyyy-mm-dd", "yyyy-MM-dd" },
            { "dd/mm/yyyy", "dd/MM/yyyy" },
            { "mm/dd/yyyy", "MM/dd/yyyy" },
            { "yyyy-mm-dd hh:mm:ss", "yyyy-MM-dd HH:mm:ss" }
        };
        }

        public object DetectAndConvert(IXLCell cell)
        {
            // Try to get the value with Excel's type detection first
            var dataType = cell.DataType;

            switch (dataType)
            {
                case XLDataType.Boolean:
                    return cell.GetValue<bool>();

                case XLDataType.Number:
                    return ConvertNumber(cell);

                case XLDataType.DateTime:
                    return ConvertDateTime(cell);

                case XLDataType.Text:
                    return ConvertText(cell);

                case XLDataType.Blank:
                    return null;

                default:
                    return cell.GetValue<string>();
            }
        }

        private object ConvertNumber(IXLCell cell)
        {
            var value = cell.GetValue<double>();

            // Check if it's actually a date (Excel stores dates as numbers)
            if (cell.Style.DateFormat != null && !string.IsNullOrEmpty(cell.Style.DateFormat.Format))
            {
                return DateTime.FromOADate(value);
            }

            // Determine if it's an integer or decimal
            if (Math.Abs(value % 1) < double.Epsilon)
            {
                // It's a whole number
                if (value >= int.MinValue && value <= int.MaxValue)
                {
                    return (int)value;
                }
                else if (value >= long.MinValue && value <= long.MaxValue)
                {
                    return (long)value;
                }
            }

            // Return as decimal for precision
            if (_options.UseDecimalForNumbers)
            {
                return (decimal)value;
            }

            return value;
        }

        private object ConvertDateTime(IXLCell cell)
        {
            try
            {
                var dateTime = cell.GetValue<DateTime>();

                if (_options.ConvertDatesToStrings)
                {
                    return dateTime.ToString(_options.DateFormat);
                }

                return dateTime;
            }
            catch
            {
                // If conversion fails, return as string
                return cell.GetValue<string>();
            }
        }

        private object ConvertText(IXLCell cell)
        {
            var text = cell.GetValue<string>();

            if (string.IsNullOrWhiteSpace(text))
                return _options.PreserveNullValues ? null : "";

            // Try to parse special types if configured
            if (_options.AutoDetectTypes)
            {
                // Check cache first
                if (_typeCache.TryGetValue(text, out var cachedType))
                {
                    return ConvertCachedType(text, cachedType);
                }

                // Try parsing as various types
                var converted = TryParseSpecialTypes(text);
                if (converted != null)
                {
                    _typeCache[text] = converted.GetType();
                    return converted;
                }
            }

            return text;
        }

        private object TryParseSpecialTypes(string text)
        {
            // Boolean
            if (bool.TryParse(text, out var boolValue))
                return boolValue;

            // Integer
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var intValue))
                return intValue;

            // Long
            if (long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var longValue))
                return longValue;

            // Decimal
            if (decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out var decimalValue))
                return _options.UseDecimalForNumbers ? decimalValue : (double)decimalValue;

            // DateTime
            if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateValue))
            {
                return _options.ConvertDatesToStrings ? text : dateValue;
            }

            // GUID
            if (Guid.TryParse(text, out var guidValue))
                return _options.PreserveGuidsAsStrings ? text : guidValue;

            // JSON
            if (_options.ParseJsonStrings && (text.StartsWith("{") || text.StartsWith("[")))
            {
                try
                {
                    return JToken.Parse(text);
                }
                catch
                {
                    // Not valid JSON
                }
            }

            return null;
        }

        private object ConvertCachedType(string text, Type type)
        {
            try
            {
                if (type == typeof(bool))
                    return bool.Parse(text);
                if (type == typeof(int))
                    return int.Parse(text, CultureInfo.InvariantCulture);
                if (type == typeof(long))
                    return long.Parse(text, CultureInfo.InvariantCulture);
                if (type == typeof(decimal))
                    return decimal.Parse(text, CultureInfo.InvariantCulture);
                if (type == typeof(double))
                    return double.Parse(text, CultureInfo.InvariantCulture);
                if (type == typeof(DateTime))
                    return DateTime.Parse(text, CultureInfo.InvariantCulture);
                if (type == typeof(Guid))
                    return Guid.Parse(text);
            }
            catch
            {
                // Conversion failed, remove from cache
                _typeCache.Remove(text);
            }

            return text;
        }
    }

    /// <summary>
    /// Analyzes data structures for intelligent grouping
    /// </summary>
    public class StructureAnalyzer
    {
        public Dictionary<string, List<Dictionary<string, object>>> GroupByStructure(List<Dictionary<string, object>> data)
        {
            var groups = new Dictionary<string, List<Dictionary<string, object>>>();

            foreach (var record in data)
            {
                var signature = GetStructureSignature(record);

                if (!groups.ContainsKey(signature))
                {
                    groups[signature] = new List<Dictionary<string, object>>();
                }

                groups[signature].Add(record);
            }

            return RenameGroupsForClarity(groups);
        }

        private string GetStructureSignature(Dictionary<string, object> record)
        {
            var signature = new StringBuilder();
            var typeCount = new Dictionary<string, int>();

            foreach (var kvp in record.OrderBy(x => x.Key))
            {
                var type = kvp.Value?.GetType().Name ?? "null";

                if (!typeCount.ContainsKey(type))
                    typeCount[type] = 0;

                typeCount[type]++;
            }

            foreach (var tc in typeCount.OrderBy(x => x.Key))
            {
                signature.Append($"{tc.Key}:{tc.Value},");
            }

            return signature.ToString().TrimEnd(',');
        }

        private Dictionary<string, List<Dictionary<string, object>>> RenameGroupsForClarity(
            Dictionary<string, List<Dictionary<string, object>>> groups)
        {
            var renamedGroups = new Dictionary<string, List<Dictionary<string, object>>>();
            int groupIndex = 1;

            foreach (var group in groups.OrderByDescending(g => g.Value.Count))
            {
                var name = $"Group{groupIndex}_Records{group.Value.Count}";
                renamedGroups[name] = group.Value;
                groupIndex++;
            }

            return renamedGroups;
        }
    }

    /// <summary>
    /// Configuration options for Excel import
    /// </summary>
    public class ExcelImportOptions
    {
        public int HeaderRow { get; set; } = 1;
        public bool SkipEmptyRows { get; set; } = true;
        public bool AutoDetectTypes { get; set; } = true;
        public bool UseDecimalForNumbers { get; set; } = true;
        public bool ConvertDatesToStrings { get; set; } = false;
        public string DateFormat { get; set; } = "yyyy-MM-dd HH:mm:ss";
        public bool PreserveNullValues { get; set; } = true;
        public bool PreserveGuidsAsStrings { get; set; } = true;
        public bool ParseJsonStrings { get; set; } = true;
        public bool EvaluateFormulas { get; set; } = true;
        public bool GenerateMissingHeaders { get; set; } = true;
        public bool IncludeSheetMetadata { get; set; } = false;
        public Formatting JsonFormatting { get; set; } = Formatting.Indented;
    }

    /// <summary>
    /// Result of Excel to JSON conversion
    /// </summary>
    public class ConversionResult
    {
        public bool Success { get; set; }
        public string JsonString { get; set; }
        public object Data { get; set; }
        public string ErrorMessage { get; set; }
        public Exception Exception { get; set; }
        public int RecordCount { get; set; }
        public int SheetCount { get; set; }
        public int GroupCount { get; set; }
        public ConversionMode ConversionMode { get; set; }
        public Dictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();
    }

    /// <summary>
    /// Conversion modes for different Excel structures
    /// </summary>
    public enum ConversionMode
    {
        Auto,      // Automatically detect best mode
        Simple,    // Simple flat structure
        Nested,    // Reconstruct nested objects from flat columns
        MultiSheet,// Multiple sheets as separate arrays
        Grouped    // Group by structure similarity
    }


}