using System;
using System.Runtime.Serialization;

namespace ExportExcel.Exceptions
{
    /// <summary>
    /// Base exception for all Excel export/import operations
    /// </summary>
    [Serializable]
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

#pragma warning disable SYSLIB0051 // Type or member is obsolete
        protected ExcelOperationException(SerializationInfo info, StreamingContext context) : base(info, context)
#pragma warning restore SYSLIB0051 // Type or member is obsolete
        {
            OperationType = info.GetString(nameof(OperationType));
            Context = info.GetValue(nameof(Context), typeof(object));
        }

        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
            info.AddValue(nameof(OperationType), OperationType);
            info.AddValue(nameof(Context), Context);
        }
    }

    /// <summary>
    /// Exception thrown when Excel export operations fail
    /// </summary>
    [Serializable]
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

        protected ExcelExportException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
            RecordCount = (int?)info.GetValue(nameof(RecordCount), typeof(int?));
            SheetName = info.GetString(nameof(SheetName));
        }

        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
            info.AddValue(nameof(RecordCount), RecordCount);
            info.AddValue(nameof(SheetName), SheetName);
        }
    }

    /// <summary>
    /// Exception thrown when Excel import operations fail
    /// </summary>
    [Serializable]
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

        protected ExcelImportException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
            WorksheetName = info.GetString(nameof(WorksheetName));
            RowNumber = (int?)info.GetValue(nameof(RowNumber), typeof(int?));
            ColumnNumber = (int?)info.GetValue(nameof(ColumnNumber), typeof(int?));
            CellReference = info.GetString(nameof(CellReference));
        }

        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
            info.AddValue(nameof(WorksheetName), WorksheetName);
            info.AddValue(nameof(RowNumber), RowNumber);
            info.AddValue(nameof(ColumnNumber), ColumnNumber);
            info.AddValue(nameof(CellReference), CellReference);
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
    [Serializable]
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

        protected JsonParsingException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
            JsonFragment = info.GetString(nameof(JsonFragment));
            LineNumber = (int?)info.GetValue(nameof(LineNumber), typeof(int?));
            Position = (int?)info.GetValue(nameof(Position), typeof(int?));
        }

        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
            info.AddValue(nameof(JsonFragment), JsonFragment);
            info.AddValue(nameof(LineNumber), LineNumber);
            info.AddValue(nameof(Position), Position);
        }
    }

    /// <summary>
    /// Exception thrown when data validation fails
    /// </summary>
    [Serializable]
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

        protected DataValidationException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
            PropertyName = info.GetString(nameof(PropertyName));
            InvalidValue = info.GetValue(nameof(InvalidValue), typeof(object));
            ValidationRules = (string[])info.GetValue(nameof(ValidationRules), typeof(string[]));
        }

        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
            info.AddValue(nameof(PropertyName), PropertyName);
            info.AddValue(nameof(InvalidValue), InvalidValue);
            info.AddValue(nameof(ValidationRules), ValidationRules);
        }
    }

    /// <summary>
    /// Exception thrown when worksheet operations fail
    /// </summary>
    [Serializable]
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

        protected WorksheetException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
            WorksheetName = info.GetString(nameof(WorksheetName));
            WorksheetOperation = info.GetString(nameof(WorksheetOperation));
        }

        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
            info.AddValue(nameof(WorksheetName), WorksheetName);
            info.AddValue(nameof(WorksheetOperation), WorksheetOperation);
        }
    }

    /// <summary>
    /// Exception thrown when structure analysis fails
    /// </summary>
    [Serializable]
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

        protected StructureAnalysisException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
            StructureType = info.GetString(nameof(StructureType));
            RecordIndex = info.GetInt32(nameof(RecordIndex));
        }

        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
            info.AddValue(nameof(StructureType), StructureType);
            info.AddValue(nameof(RecordIndex), RecordIndex);
        }
    }

}
