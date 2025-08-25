using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportExcel.Interfaces
{
    /// <summary>
    /// Interface for async Excel export operations
    /// </summary>
    public interface IAsyncExcelExporter : IExcelExporter
    {
        /// <summary>
        /// Asynchronously exports data to Excel
        /// </summary>
        Task<byte[]> ExportToExcelAsync(List<Dictionary<string, object>> dataList, string sheetName = "Data");

        /// <summary>
        /// Asynchronously exports JSON to Excel
        /// </summary>
        Task<byte[]> ExportJsonToExcelAsync(string jsonString, string sheetName = "Data");

        /// <summary>
        /// Asynchronously exports grouped JSON to Excel
        /// </summary>
        Task<byte[]> ExportGroupedJsonToExcelAsync(string jsonString, string groupByField = "category");

        /// <summary>
        /// Asynchronously exports flattened JSON to Excel
        /// </summary>
        Task<byte[]> ExportFlattenedJsonToExcelAsync(string jsonString);

        /// <summary>
        /// Asynchronously exports flattened grouped JSON to Excel
        /// </summary>
        Task<byte[]> ExportFlattenedGroupedJsonToExcelAsync(string jsonString, string groupByField = "type");
    }

}
