using Newtonsoft.Json;

namespace ExportExcel.Interfaces
{
    /// <summary>
    /// Interface for export configuration options
    /// </summary>
    public interface IExcelExportOptions
    {
        /// <summary>
        /// Default worksheet name
        /// </summary>
        string DefaultSheetName { get; set; }

        /// <summary>
        /// Whether to include filtering on headers
        /// </summary>
        bool EnableAutoFilter { get; set; }

        /// <summary>
        /// Whether to freeze header rows
        /// </summary>
        bool FreezeHeaderRows { get; set; }

        /// <summary>
        /// Whether to auto-adjust column widths
        /// </summary>
        bool AutoAdjustColumnWidth { get; set; }

        /// <summary>
        /// Maximum number of records per sheet
        /// </summary>
        int MaxRecordsPerSheet { get; set; }

        /// <summary>
        /// Date format string for date cells
        /// </summary>
        string DateFormat { get; set; }

        /// <summary>
        /// Number format for numeric cells
        /// </summary>
        string NumberFormat { get; set; }

        /// <summary>
        /// Whether to create summary sheets
        /// </summary>
        bool CreateSummarySheet { get; set; }

        /// <summary>
        /// JSON formatting for JSON exports
        /// </summary>
        Formatting JsonFormatting { get; set; }
    }
}
