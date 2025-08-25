using ClosedXML.Excel;
using System.Collections.Generic;

namespace ExportExcel.Interfaces
{
    /// <summary>
    /// Interface for Excel worksheet operations
    /// </summary>
    public interface IWorksheetManager
    {
        /// <summary>
        /// Creates a worksheet with the provided data and styling
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="sheetName">Name of the sheet</param>
        /// <param name="data">Data to populate</param>
        /// <param name="headers">Column headers</param>
        void CreateWorksheet(XLWorkbook workbook, string sheetName, List<Dictionary<string, object>> data, IEnumerable<string> headers);

        /// <summary>
        /// Creates a summary worksheet with statistics
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="summaryData">Summary information</param>
        /// <param name="title">Summary title</param>
        void CreateSummarySheet(XLWorkbook workbook, Dictionary<string, object> summaryData, string title = "Summary");

        /// <summary>
        /// Applies consistent styling to header range
        /// </summary>
        /// <param name="range">Range to style</param>
        void StyleHeaderRange(IXLRange range);

        /// <summary>
        /// Applies consistent styling to data range
        /// </summary>
        /// <param name="range">Range to style</param>
        void StyleDataRange(IXLRange range);

        /// <summary>
        /// Sets cell value with intelligent type detection
        /// </summary>
        /// <param name="cell">Target cell</param>
        /// <param name="value">Value to set</param>
        void SetCellValueWithType(IXLCell cell, object value);

        /// <summary>
        /// Sanitizes sheet names to comply with Excel naming rules
        /// </summary>
        /// <param name="name">Original name</param>
        /// <returns>Sanitized name</returns>
        string SanitizeSheetName(string name);
    }

}
