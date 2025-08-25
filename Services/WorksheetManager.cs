using ClosedXML.Excel;
using ExportExcel.Exceptions;
using ExportExcel.Interfaces;
using ExportExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportExcel.Services
{
    /// <summary>
    /// Excel worksheet management service
    /// </summary>
    public class WorksheetManager : IWorksheetManager
    {
        private readonly ITypeDetector _typeDetector;
        private readonly Dictionary<string, Type> _typeCache;
        private readonly ExcelExportOptions _options;

        public WorksheetManager(ITypeDetector typeDetector = null, ExcelExportOptions options = null)
        {
            _typeDetector = typeDetector;
            _typeCache = new Dictionary<string, Type>();
            _options = options ?? new ExcelExportOptions();
        }

        public void CreateWorksheet(XLWorkbook workbook, string sheetName, List<Dictionary<string, object>> data, IEnumerable<string> headers)
        {
            try
            {
                if (workbook == null) throw new ArgumentNullException(nameof(workbook));
                if (string.IsNullOrWhiteSpace(sheetName)) sheetName = _options.DefaultSheetName;

                var sanitizedName = SanitizeSheetName(sheetName);
                var worksheet = workbook.Worksheets.Add(sanitizedName);
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

                if (_options.EnableAutoFilter)
                {
                    headerRange.SetAutoFilter();
                }

                if (_options.FreezeHeaderRows)
                {
                    worksheet.SheetView.FreezeRows(1);
                }

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

                if (_options.AutoAdjustColumnWidth)
                {
                    worksheet.Columns().AdjustToContents();
                }
            }
            catch (Exception ex)
            {
                throw new WorksheetException(
                    $"Failed to create worksheet '{sheetName}': {ex.Message}",
                    sheetName,
                    ex,
                    "CreateWorksheet");
            }
        }

        public void CreateSummarySheet(XLWorkbook workbook, Dictionary<string, object> summaryData, string title = "Summary")
        {
            try
            {
                if (workbook == null) throw new ArgumentNullException(nameof(workbook));
                if (summaryData == null) throw new ArgumentNullException(nameof(summaryData));

                var summarySheet = workbook.Worksheets.Add(title);

                // Create title
                var titleCell = summarySheet.Cell(1, 1);
                titleCell.Value = title;
                titleCell.Style.Font.Bold = true;
                titleCell.Style.Font.FontSize = 16;
                titleCell.Style.Fill.BackgroundColor = XLColor.LightBlue;
                titleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                summarySheet.Range(1, 1, 1, 2).Merge();

                // Add summary data
                int row = 3;
                foreach (var kvp in summaryData)
                {
                    summarySheet.Cell(row, 1).Value = kvp.Key;
                    summarySheet.Cell(row, 1).Style.Font.Bold = true;
                    summarySheet.Cell(row, 2).Value = kvp.Value?.ToString() ?? "";
                    row++;
                }

                // Style the data range
                var dataRange = summarySheet.Range(3, 1, row - 1, 2);
                StyleDataRange(dataRange);

                summarySheet.Columns().AdjustToContents();
            }
            catch (Exception ex)
            {
                throw new WorksheetException(
                    $"Failed to create summary sheet '{title}': {ex.Message}",
                    title,
                    ex,
                    "CreateSummarySheet");
            }
        }

        public void StyleHeaderRange(IXLRange range)
        {
            if (range == null) return;

            range.Style.Font.Bold = true;
            range.Style.Fill.BackgroundColor = XLColor.LightGray;
            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            range.Style.Alignment.WrapText = true;
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        public void StyleDataRange(IXLRange range)
        {
            if (range == null) return;

            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            range.Style.Alignment.WrapText = true;
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        public void SetCellValueWithType(IXLCell cell, object value)
        {
            if (cell == null) return;

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
                    cell.Style.DateFormat.Format = _options.DateFormat;
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
                cell.Style.DateFormat.Format = _options.DateFormat;
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

        public string SanitizeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return _options.DefaultSheetName;

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

        private void SetCellValueFromCachedType(IXLCell cell, string str, Type type)
        {
            try
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
                    cell.Style.DateFormat.Format = _options.DateFormat;
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
            catch
            {
                // Conversion failed, remove from cache and use as string
                _typeCache.Remove(str);
                cell.Value = str;
            }
        }
    }

}
