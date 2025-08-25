using ExcelExport;
using System.Collections.Generic;
using System.IO;

namespace ExportExcel.Interfaces
{
    /// <summary>
    /// Core interface for Excel to JSON conversion operations
    /// </summary>
    public interface IExcelImporter
    {
        /// <summary>
        /// Converts Excel file bytes to list of data with automatic structure detection
        /// </summary>
        /// <param name="excelBytes">Excel file bytes</param>
        /// <param name="mode">Conversion mode</param>
        /// <returns>List of dictionaries representing the data</returns>
        List<Dictionary<string, object>> ConvertToData(byte[] excelBytes, ConversionMode mode = ConversionMode.Auto);

        /// <summary>
        /// Converts Excel file from stream to list of data
        /// </summary>
        /// <param name="stream">Excel file stream</param>
        /// <param name="mode">Conversion mode</param>
        /// <returns>List of dictionaries representing the data</returns>
        List<Dictionary<string, object>> ConvertFromStream(Stream stream, ConversionMode mode = ConversionMode.Auto);

        /// <summary>
        /// Legacy method - returns ConversionResult for backward compatibility
        /// </summary>
        /// <param name="excelBytes">Excel file bytes</param>
        /// <param name="mode">Conversion mode</param>
        /// <returns>Conversion result with metadata</returns>
        ConversionResult ConvertToJson(byte[] excelBytes, ConversionMode mode = ConversionMode.Auto);
    }

}
