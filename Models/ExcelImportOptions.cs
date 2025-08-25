using ExportExcel.Interfaces;

namespace ExportExcel.Models
{
    /// <summary>
    /// Configuration for import operations
    /// </summary>
    public class ExcelImportOptions : IExcelImportOptions
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
        public ValidationLevel ValidationLevel { get; set; } = ValidationLevel.Standard;
        public ErrorHandlingStrategy ErrorHandling { get; set; } = ErrorHandlingStrategy.ThrowImmediately;
        public int MaxRowsPerSheet { get; set; } = 1048576;
        public bool TrimWhitespace { get; set; } = true;
    }

}
