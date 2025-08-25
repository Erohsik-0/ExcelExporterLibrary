

namespace ExportExcel.Interfaces
{
    /// <summary>
    /// Interface for import configuration options
    /// </summary>
    public interface IExcelImportOptions
    {
        /// <summary>
        /// Header row number (1-based)
        /// </summary>
        int HeaderRow { get; set; }

        /// <summary>
        /// Whether to skip empty rows
        /// </summary>
        bool SkipEmptyRows { get; set; }

        /// <summary>
        /// Whether to auto-detect data types
        /// </summary>
        bool AutoDetectTypes { get; set; }

        /// <summary>
        /// Whether to use decimal for numbers
        /// </summary>
        bool UseDecimalForNumbers { get; set; }

        /// <summary>
        /// Whether to convert dates to strings
        /// </summary>
        bool ConvertDatesToStrings { get; set; }

        /// <summary>
        /// Date format for string conversion
        /// </summary>
        string DateFormat { get; set; }

        /// <summary>
        /// Whether to preserve null values
        /// </summary>
        bool PreserveNullValues { get; set; }

        /// <summary>
        /// Whether to preserve GUIDs as strings
        /// </summary>
        bool PreserveGuidsAsStrings { get; set; }

        /// <summary>
        /// Whether to parse JSON strings
        /// </summary>
        bool ParseJsonStrings { get; set; }

        /// <summary>
        /// Whether to evaluate formulas
        /// </summary>
        bool EvaluateFormulas { get; set; }

        /// <summary>
        /// Whether to generate missing headers
        /// </summary>
        bool GenerateMissingHeaders { get; set; }

        /// <summary>
        /// Whether to include sheet metadata
        /// </summary>
        bool IncludeSheetMetadata { get; set; }
    }

}
