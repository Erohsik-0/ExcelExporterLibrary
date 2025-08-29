using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using ConversionResult = ExportExcel.Models.ConversionResult;
using ConversionMode = ExportExcel.Models.ConversionMode;

namespace ExportExcel.Interfaces
{
    /// <summary>
    /// Interface for async Excel import operations
    /// </summary>
    public interface IAsyncExcelImporter : IExcelImporter
    {
        /// <summary>
        /// Asynchronously converts Excel file bytes to list of data
        /// </summary>
        Task<List<Dictionary<string, object>>> ConvertToDataAsync(byte[] excelBytes, ConversionMode mode = ConversionMode.Auto);

        /// <summary>
        /// Asynchronously converts Excel file from stream to list of data
        /// </summary>
        Task<List<Dictionary<string, object>>> ConvertFromStreamAsync(Stream stream, ConversionMode mode = ConversionMode.Auto);

        /// <summary>
        /// Asynchronously converts Excel file from file path to list of data
        /// </summary>
        Task<List<Dictionary<string, object>>> ConvertFromFileAsync(string filePath, ConversionMode mode = ConversionMode.Auto);

        /// <summary>
        /// Asynchronously converts to JSON result
        /// </summary>
        Task<ConversionResult> ConvertToJsonAsync(byte[] excelBytes, ConversionMode mode = ConversionMode.Auto);
    }

}
