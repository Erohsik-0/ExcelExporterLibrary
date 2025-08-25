using ClosedXML.Excel;
using ExportExcel.Exceptions;
using ExportExcel.Interfaces;
using ExportExcel.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExportExcel.Services
{
    /// <summary>
    /// Core Excel export service implementation
    /// </summary>
    public class ExcelExporter : IExcelExporter, IAsyncExcelExporter
    {
        private readonly IWorksheetManager _worksheetManager;
        private readonly IJsonFlattener _jsonFlattener;
        private readonly IStructureAnalyzer _structureAnalyzer;
        private readonly IDataValidator _validator;
        private readonly ExcelExportOptions _options;

        public ExcelExporter(
            IWorksheetManager worksheetManager,
            IJsonFlattener jsonFlattener,
            IStructureAnalyzer structureAnalyzer,
            IDataValidator validator = null,
            ExcelExportOptions options = null)
        {
            _worksheetManager = worksheetManager ?? throw new ArgumentNullException(nameof(worksheetManager));
            _jsonFlattener = jsonFlattener ?? throw new ArgumentNullException(nameof(jsonFlattener));
            _structureAnalyzer = structureAnalyzer ?? throw new ArgumentNullException(nameof(structureAnalyzer));
            _validator = validator;
            _options = options ?? new ExcelExportOptions();
        }

        #region Synchronous Methods

        public byte[] ExportToExcel(List<Dictionary<string, object>> dataList, string sheetName = "Data")
        {
            try
            {
                // Validate data if validator is configured
                if (_validator != null)
                {
                    var validationResult = _validator.ValidateData(dataList);
                    if (!validationResult.IsValid)
                    {
                        throw new DataValidationException(
                            $"Data validation failed: {string.Join(", ", validationResult.Errors)}");
                    }
                }

                // Handle empty data gracefully
                if (dataList == null || dataList.Count == 0)
                {
                    return CreateEmptyWorkbook(sheetName ?? _options.DefaultSheetName);
                }

                using var workbook = new XLWorkbook();
                var sanitizedSheetName = _worksheetManager.SanitizeSheetName(sheetName ?? _options.DefaultSheetName);
                var headers = GetAllUniqueHeaders(dataList);

                _worksheetManager.CreateWorksheet(workbook, sanitizedSheetName, dataList, headers);

                if (_options.CreateSummarySheet && dataList.Count > 100)
                {
                    CreateDataSummarySheet(workbook, dataList, headers);
                }

                return SaveWorkbook(workbook);
            }
            catch (Exception ex) when (!(ex is ExcelOperationException))
            {
                throw new ExcelExportException(
                    $"Failed to export data to Excel: {ex.Message}",
                    sheetName ?? _options.DefaultSheetName,
                    ex,
                    dataList?.Count);
            }
        }

        public byte[] ExportJsonToExcel(string jsonString, string sheetName = "Data")
        {
            try
            {
                if (string.IsNullOrWhiteSpace(jsonString))
                {
                    return CreateEmptyWorkbook(sheetName ?? _options.DefaultSheetName);
                }

                var dataList = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(jsonString);
                return ExportToExcel(dataList, sheetName);
            }
            catch (JsonException ex)
            {
                throw new JsonParsingException(
                    $"Failed to parse JSON: {ex.Message}",
                    jsonString?.Substring(0, Math.Min(jsonString.Length, 100)),
                    ex);
            }
        }

        public byte[] ExportGroupedJsonToExcel(string jsonString, string groupByField = "category")
        {
            try
            {
                if (string.IsNullOrWhiteSpace(jsonString))
                {
                    return CreateEmptyWorkbook("Data");
                }

                var flatData = _jsonFlattener.FlattenJson(jsonString);
                if (flatData.Count == 0)
                {
                    return CreateEmptyWorkbook("Data");
                }

                var groups = _structureAnalyzer.GroupByStructure(flatData);

                using var workbook = new XLWorkbook();

                foreach (var group in groups.OrderByDescending(g => g.Value.Count))
                {
                    var sanitizedName = _worksheetManager.SanitizeSheetName(group.Key);
                    var headers = GetAllUniqueHeaders(group.Value);
                    _worksheetManager.CreateWorksheet(workbook, sanitizedName, group.Value, headers);
                }

                if (_options.CreateSummarySheet)
                {
                    CreateGroupSummarySheet(workbook, groups);
                }

                return SaveWorkbook(workbook);
            }
            catch (Exception ex) when (!(ex is ExcelOperationException))
            {
                throw new ExcelExportException(
                    $"Failed to export grouped JSON to Excel: {ex.Message}",
                    "Grouped",
                    ex);
            }
        }

        public byte[] ExportFlattenedJsonToExcel(string jsonString)
        {
            try
            {
                var flatData = _jsonFlattener.FlattenJson(jsonString);
                return ExportToExcel(flatData, "Flattened_Data");
            }
            catch (Exception ex) when (!(ex is ExcelOperationException))
            {
                throw new ExcelExportException(
                    $"Failed to export flattened JSON to Excel: {ex.Message}",
                    "Flattened_Data",
                    ex);
            }
        }

        public byte[] ExportFlattenedGroupedJsonToExcel(string jsonString, string groupByField = "type")
        {
            try
            {
                var flatData = _jsonFlattener.FlattenJson(jsonString);
                var groups = _structureAnalyzer.GroupByStructure(flatData);

                using var workbook = new XLWorkbook();

                foreach (var group in groups.OrderByDescending(g => g.Value.Count))
                {
                    var sanitizedName = _worksheetManager.SanitizeSheetName($"Flat_{group.Key}");
                    var headers = GetAllUniqueHeaders(group.Value);
                    _worksheetManager.CreateWorksheet(workbook, sanitizedName, group.Value, headers);
                }

                if (_options.CreateSummarySheet)
                {
                    CreateGroupSummarySheet(workbook, groups);
                }

                return SaveWorkbook(workbook);
            }
            catch (Exception ex) when (!(ex is ExcelOperationException))
            {
                throw new ExcelExportException(
                    $"Failed to export flattened grouped JSON to Excel: {ex.Message}",
                    "FlattenedGrouped",
                    ex);
            }
        }

        #endregion

        #region Async Methods

        public async Task<byte[]> ExportToExcelAsync(List<Dictionary<string, object>> dataList, string sheetName = "Data")
        {
            return await Task.Run(() => ExportToExcel(dataList, sheetName));
        }

        public async Task<byte[]> ExportJsonToExcelAsync(string jsonString, string sheetName = "Data")
        {
            return await Task.Run(() => ExportJsonToExcel(jsonString, sheetName));
        }

        public async Task<byte[]> ExportGroupedJsonToExcelAsync(string jsonString, string groupByField = "category")
        {
            return await Task.Run(() => ExportGroupedJsonToExcel(jsonString, groupByField));
        }

        public async Task<byte[]> ExportFlattenedJsonToExcelAsync(string jsonString)
        {
            return await Task.Run(() => ExportFlattenedJsonToExcel(jsonString));
        }

        public async Task<byte[]> ExportFlattenedGroupedJsonToExcelAsync(string jsonString, string groupByField = "type")
        {
            return await Task.Run(() => ExportFlattenedGroupedJsonToExcel(jsonString, groupByField));
        }

        #endregion

        #region Helper Methods

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

        private byte[] CreateEmptyWorkbook(string sheetName)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add(sheetName);
            worksheet.Cell(1, 1).Value = "No data available";
            return SaveWorkbook(workbook);
        }

        private void CreateDataSummarySheet(XLWorkbook workbook, List<Dictionary<string, object>> data, string[] headers)
        {
            var summaryData = new Dictionary<string, object>
            {
                ["TotalRecords"] = data.Count,
                ["TotalColumns"] = headers.Length,
                ["GeneratedAt"] = DateTime.Now.ToString(_options.DateFormat),
                ["Headers"] = string.Join(", ", headers.Take(10)) + (headers.Length > 10 ? "..." : "")
            };

            _worksheetManager.CreateSummarySheet(workbook, summaryData, "Data Summary");
        }

        private void CreateGroupSummarySheet(XLWorkbook workbook, Dictionary<string, List<Dictionary<string, object>>> groups)
        {
            var summaryData = new Dictionary<string, object>
            {
                ["TotalGroups"] = groups.Count,
                ["TotalRecords"] = groups.Values.Sum(g => g.Count),
                ["GeneratedAt"] = DateTime.Now.ToString(_options.DateFormat),
                ["LargestGroup"] = groups.OrderByDescending(g => g.Value.Count).First().Key,
                ["GroupSizes"] = string.Join(", ", groups.Select(g => $"{g.Key}: {g.Value.Count}").Take(5))
            };

            _worksheetManager.CreateSummarySheet(workbook, summaryData, "Groups Summary");
        }

        private byte[] SaveWorkbook(XLWorkbook workbook)
        {
            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return stream.ToArray();
        }

        #endregion
    }

}
