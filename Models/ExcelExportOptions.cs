using ExportExcel.Interfaces;
using Newtonsoft.Json;

namespace ExportExcel.Models
{
    /// <summary>
    /// Configuration for export operations
    /// </summary>
    public class ExcelExportOptions : IExcelExportOptions
    {
        public string DefaultSheetName { get; set; } = "Data";
        public bool EnableAutoFilter { get; set; } = true;
        public bool FreezeHeaderRows { get; set; } = true;
        public bool AutoAdjustColumnWidth { get; set; } = true;
        public int MaxRecordsPerSheet { get; set; } = 1048576; // Excel row limit
        public string DateFormat { get; set; } = "yyyy-MM-dd HH:mm:ss";
        public string NumberFormat { get; set; } = "#,##0.00";
        public bool CreateSummarySheet { get; set; } = true;
        public Formatting JsonFormatting { get; set; } = Formatting.Indented;
        public ValidationLevel ValidationLevel { get; set; } = ValidationLevel.Standard;
        public ErrorHandlingStrategy ErrorHandling { get; set; } = ErrorHandlingStrategy.ThrowImmediately;
        public bool IncludeMetadata { get; set; } = true;
        public bool OptimizeForLargeData { get; set; } = false;
    }

}
