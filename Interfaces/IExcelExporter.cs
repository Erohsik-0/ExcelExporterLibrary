using System.Collections.Generic;

namespace ExportExcel.Interfaces
{
    /// <summary>
    /// Core interface for Excel export operations
    /// </summary>
    public interface IExcelExporter
    {
        /// <summary>
        /// Exports a list of dictionaries to Excel format with proper styling and data types
        /// </summary>
        /// <param name="dataList">The data to export as key-value pairs</param>
        /// <param name="sheetName">Name of the worksheet (defaults to "Data")</param>
        /// <returns>Excel file as byte array</returns>
        byte[] ExportToExcel(List<Dictionary<string, object>> dataList, string sheetName = "Data");

        /// <summary>
        /// Converts JSON string to Excel by deserializing to dictionary list
        /// </summary>
        /// <param name="jsonString">JSON string to convert</param>
        /// <param name="sheetName">Name of the worksheet</param>
        /// <returns>Excel file as byte array</returns>
        byte[] ExportJsonToExcel(string jsonString, string sheetName = "Data");

        /// <summary>
        /// Creates multiple Excel sheets grouped by nested content structure
        /// </summary>
        /// <param name="jsonString">JSON data to process</param>
        /// <param name="groupByField">Optional field to use for additional grouping logic</param>
        /// <returns>Excel file as byte array</returns>
        byte[] ExportGroupedJsonToExcel(string jsonString, string groupByField = "category");

        /// <summary>
        /// Flattens nested JSON structures into a single-level Excel sheet
        /// </summary>
        /// <param name="jsonString">JSON string with nested structures</param>
        /// <returns>Excel file as byte array</returns>
        byte[] ExportFlattenedJsonToExcel(string jsonString);

        /// <summary>
        /// Combines flattening with grouping by nested content structure
        /// </summary>
        /// <param name="jsonString">JSON string to process</param>
        /// <param name="groupByField">Field to group by</param>
        /// <returns>Excel file as byte array</returns>
        byte[] ExportFlattenedGroupedJsonToExcel(string jsonString, string groupByField = "type");
    }

}
