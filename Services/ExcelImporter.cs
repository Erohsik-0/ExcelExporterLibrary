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
    /// Excel to data/JSON conversion service implementation
    /// </summary>
    public class ExcelImporter : IExcelImporter, IAsyncExcelImporter
    {
        private readonly ITypeDetector _typeDetector;
        private readonly IStructureAnalyzer _structureAnalyzer;
        private readonly IJsonFlattener _jsonFlattener;
        private readonly IDataValidator _validator;
        private readonly ExcelImportOptions _options;

        public ExcelImporter(
            ITypeDetector typeDetector,
            IStructureAnalyzer structureAnalyzer,
            IJsonFlattener jsonFlattener,
            IDataValidator validator = null,
            ExcelImportOptions options = null)
        {
            _typeDetector = typeDetector ?? throw new ArgumentNullException(nameof(typeDetector));
            _structureAnalyzer = structureAnalyzer ?? throw new ArgumentNullException(nameof(structureAnalyzer));
            _jsonFlattener = jsonFlattener ?? throw new ArgumentNullException(nameof(jsonFlattener));
            _validator = validator;
            _options = options ?? new ExcelImportOptions();
        }

        #region Synchronous Methods

        public List<Dictionary<string, object>> ConvertToData(byte[] excelBytes, ConversionMode mode = ConversionMode.Auto)
        {
            try
            {
                using var stream = new MemoryStream(excelBytes);
                using var workbook = new XLWorkbook(stream);

                if (workbook.Worksheets.Count == 0)
                {
                    throw new ExcelImportException("Excel file contains no worksheets", "Unknown");
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
            catch (Exception ex) when (!(ex is ExcelOperationException))
            {
                throw new ExcelImportException($"Conversion failed: {ex.Message}", "Unknown", ex);
            }
        }

        public List<Dictionary<string, object>> ConvertFromStream(Stream stream, ConversionMode mode = ConversionMode.Auto)
        {
            using var memoryStream = new MemoryStream();
            stream.CopyTo(memoryStream);
            return ConvertToData(memoryStream.ToArray(), mode);
        }

        public ConversionResult ConvertToJson(byte[] excelBytes, ConversionMode mode = ConversionMode.Auto)
        {
            var result = new ConversionResult();

            try
            {
                var data = ConvertToData(excelBytes, mode);

                result.Success = true;
                result.Data = data;
                result.JsonString = JsonConvert.SerializeObject(data);
                result.RecordCount = data.Count;
                result.ConversionMode = mode;

                // Add metadata
                result.AddMetadata("ProcessedAt", DateTime.UtcNow);
                result.AddMetadata("Options", _options);
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

        #region Async Methods

        public async Task<List<Dictionary<string, object>>> ConvertToDataAsync(byte[] excelBytes, ConversionMode mode = ConversionMode.Auto)
        {
            return await Task.Run(() => ConvertToData(excelBytes, mode));
        }

        public async Task<List<Dictionary<string, object>>> ConvertFromStreamAsync(Stream stream, ConversionMode mode = ConversionMode.Auto)
        {
            return await Task.Run(() => ConvertFromStream(stream, mode));
        }

        public async Task<List<Dictionary<string, object>>> ConvertFromFileAsync(string filePath, ConversionMode mode = ConversionMode.Auto)
        {
            var bytes = await File.ReadAllBytesAsync(filePath);
            return await ConvertToDataAsync(bytes, mode);
        }

        public async Task<ConversionResult> ConvertToJsonAsync(byte[] excelBytes, ConversionMode mode = ConversionMode.Auto)
        {
            return await Task.Run(() => ConvertToJson(excelBytes, mode));
        }

        #endregion

        #region Conversion Mode Methods

        private List<Dictionary<string, object>> ConvertSimple(XLWorkbook workbook)
        {
            var worksheet = workbook.Worksheets.First();
            return ExtractSheetData(worksheet);
        }

        private List<Dictionary<string, object>> ConvertNested(XLWorkbook workbook)
        {
            var worksheet = workbook.Worksheets.First();
            var flatData = ExtractSheetData(worksheet);
            return _jsonFlattener.ReconstructNestedStructure(flatData);
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

            // For grouped mode, add group information
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

        #region Data Extraction Methods

        private List<Dictionary<string, object>> ExtractSheetData(IXLWorksheet worksheet)
        {
            var data = new List<Dictionary<string, object>>();
            var usedRange = worksheet.RangeUsed();

            if (usedRange == null)
                return data;

            try
            {
                // Extract headers
                var headers = ExtractHeaders(worksheet, usedRange);
                if (headers.Count == 0)
                    return data;

                // Validate headers if validator is configured
                if (_validator != null)
                {
                    var headerValidation = _validator.ValidateHeaders(headers);
                    if (!headerValidation.IsValid && _options.ValidationLevel != ValidationLevel.None)
                    {
                        throw new DataValidationException(
                            $"Header validation failed: {string.Join(", ", headerValidation.Errors)}");
                    }
                }

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
                        // Validate record if validator is configured
                        if (_validator != null && _options.ValidationLevel != ValidationLevel.None)
                        {
                            var recordValidation = _validator.ValidateRecord(rowData, data.Count);
                            if (!recordValidation.IsValid)
                            {
                                if (_options.ErrorHandling == ErrorHandlingStrategy.ThrowImmediately)
                                {
                                    throw new DataValidationException(
                                        $"Record validation failed at row {row}: {string.Join(", ", recordValidation.Errors)}");
                                }
                                else if (_options.ErrorHandling == ErrorHandlingStrategy.SkipInvalidData)
                                {
                                    continue; // Skip this record
                                }
                            }
                        }

                        data.Add(rowData);
                    }
                }

                return data;
            }
            catch (Exception ex) when (!(ex is ExcelOperationException))
            {
                throw new ExcelImportException(
                    $"Failed to extract data from worksheet '{worksheet.Name}': {ex.Message}",
                    worksheet.Name,
                    ex);
            }
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
                if (headers.Any(h => _jsonFlattener.IsNestedField(h)))
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
    }

}
