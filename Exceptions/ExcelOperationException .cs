using System;

namespace ExportExcel.Exceptions
{
    /// <summary>
    /// Base exception for all Excel export/import operations
    /// </summary>
    public class ExcelOperationException : Exception
    {
        public string OperationType { get; }
        public object Context { get; }

        public ExcelOperationException() : base() { }

        public ExcelOperationException(string message) : base(message) { }

        public ExcelOperationException(string message, Exception innerException) : base(message, innerException) { }

        public ExcelOperationException(string message, string operationType, object context = null)
            : base(message)
        {
            OperationType = operationType;
            Context = context;
        }

        public ExcelOperationException(string message, string operationType, Exception innerException, object context = null)
            : base(message, innerException)
        {
            OperationType = operationType;
            Context = context;
        }
    }

    /// <summary>
    /// Exception thrown when Excel export operations fail
    /// </summary>
    public class ExcelExportException : ExcelOperationException
    {
        public int? RecordCount { get; }
        public string SheetName { get; }

        public ExcelExportException() : base() { }

        public ExcelExportException(string message) : base(message) { }

        public ExcelExportException(string message, Exception innerException) : base(message, innerException) { }

        public ExcelExportException(string message, string sheetName, int? recordCount = null)
            : base(message, "Export")
        {
            SheetName = sheetName;
            RecordCount = recordCount;
        }

        public ExcelExportException(string message, string sheetName, Exception innerException, int? recordCount = null)
            : base(message, "Export", innerException)
        {
            SheetName = sheetName;
            RecordCount = recordCount;
        }
    }

    /// <summary>
    /// Exception thrown when Excel import operations fail
    /// </summary>
    public class ExcelImportException : ExcelOperationException
    {
        public string WorksheetName { get; }
        public int? RowNumber { get; }
        public int? ColumnNumber { get; }
        public string CellReference { get; }

        public ExcelImportException() : base() { }

        public ExcelImportException(string message) : base(message) { }

        public ExcelImportException(string message, Exception innerException) : base(message, innerException) { }

        public ExcelImportException(string message, string worksheetName, int? rowNumber = null, int? columnNumber = null)
            : base(message, "Import")
        {
            WorksheetName = worksheetName;
            RowNumber = rowNumber;
            ColumnNumber = columnNumber;
            CellReference = (rowNumber.HasValue && columnNumber.HasValue)
                ? $"{GetColumnName(columnNumber.Value)}{rowNumber.Value}"
                : null;
        }

        public ExcelImportException(string message, string worksheetName, Exception innerException,
            int? rowNumber = null, int? columnNumber = null)
            : base(message, "Import", innerException)
        {
            WorksheetName = worksheetName;
            RowNumber = rowNumber;
            ColumnNumber = columnNumber;
            CellReference = (rowNumber.HasValue && columnNumber.HasValue)
                ? $"{GetColumnName(columnNumber.Value)}{rowNumber.Value}"
                : null;
        }

        /// <summary>
        /// Converts column number to Excel column name (A, B, C, ... AA, AB, etc.)
        /// </summary>
        private static string GetColumnName(int columnNumber)
        {
            string columnName = "";
            while (columnNumber > 0)
            {
                columnNumber--; // Adjust for 1-based indexing
                columnName = (char)('A' + columnNumber % 26) + columnName;
                columnNumber /= 26;
            }
            return columnName;
        }
    }

    /// <summary>
    /// Exception thrown when JSON parsing fails during Excel operations
    /// </summary>
    public class JsonParsingException : ExcelOperationException
    {
        public string JsonFragment { get; }
        public int? LineNumber { get; }
        public int? Position { get; }

        public JsonParsingException() : base() { }

        public JsonParsingException(string message) : base(message) { }

        public JsonParsingException(string message, Exception innerException) : base(message, innerException) { }

        public JsonParsingException(string message, string jsonFragment, int? lineNumber = null, int? position = null)
            : base(message, "JsonParsing")
        {
            JsonFragment = jsonFragment;
            LineNumber = lineNumber;
            Position = position;
        }

        public JsonParsingException(string message, string jsonFragment, Exception innerException,
            int? lineNumber = null, int? position = null)
            : base(message, "JsonParsing", innerException)
        {
            JsonFragment = jsonFragment;
            LineNumber = lineNumber;
            Position = position;
        }
    }

    /// <summary>
    /// Exception thrown when data validation fails
    /// </summary>
    public class DataValidationException : ExcelOperationException
    {
        public string PropertyName { get; }
        public object InvalidValue { get; }
        public string[] ValidationRules { get; }

        public DataValidationException() : base() { }

        public DataValidationException(string message) : base(message) { }

        public DataValidationException(string message, Exception innerException) : base(message, innerException) { }

        public DataValidationException(string message, string propertyName, object invalidValue = null, string[] validationRules = null)
            : base(message, "DataValidation")
        {
            PropertyName = propertyName;
            InvalidValue = invalidValue;
            ValidationRules = validationRules;
        }

        public DataValidationException(string message, string propertyName, Exception innerException,
            object invalidValue = null, string[] validationRules = null)
            : base(message, "DataValidation", innerException)
        {
            PropertyName = propertyName;
            InvalidValue = invalidValue;
            ValidationRules = validationRules;
        }
    }

    /// <summary>
    /// Exception thrown when worksheet operations fail
    /// </summary>
    public class WorksheetException : ExcelOperationException
    {
        public string WorksheetName { get; }
        public string WorksheetOperation { get; }

        public WorksheetException() : base() { }

        public WorksheetException(string message) : base(message) { }

        public WorksheetException(string message, Exception innerException) : base(message, innerException) { }

        public WorksheetException(string message, string worksheetName, string operation = null)
            : base(message, "Worksheet")
        {
            WorksheetName = worksheetName;
            WorksheetOperation = operation;
        }

        public WorksheetException(string message, string worksheetName, Exception innerException, string operation = null)
            : base(message, "Worksheet", innerException)
        {
            WorksheetName = worksheetName;
            WorksheetOperation = operation;
        }
    }

    /// <summary>
    /// Exception thrown when structure analysis fails
    /// </summary>
    public class StructureAnalysisException : ExcelOperationException
    {
        public string StructureType { get; }
        public int RecordIndex { get; }

        public StructureAnalysisException() : base() { }

        public StructureAnalysisException(string message) : base(message) { }

        public StructureAnalysisException(string message, Exception innerException) : base(message, innerException) { }

        public StructureAnalysisException(string message, string structureType, int recordIndex = -1)
            : base(message, "StructureAnalysis")
        {
            StructureType = structureType;
            RecordIndex = recordIndex;
        }

        public StructureAnalysisException(string message, string structureType, Exception innerException, int recordIndex = -1)
            : base(message, "StructureAnalysis", innerException)
        {
            StructureType = structureType;
            RecordIndex = recordIndex;
        }
    }
}