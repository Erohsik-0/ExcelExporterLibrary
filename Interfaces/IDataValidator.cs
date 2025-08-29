using ExportExcel.Models;
using System.Collections.Generic;


namespace ExportExcel.Interfaces
{
    /// <summary>
    /// Interface for data validation operations
    /// </summary>
    public interface IDataValidator
    {
        /// <summary>
        /// Validates data before export/import operations
        /// </summary>
        /// <param name="data">Data to validate</param>
        /// <returns>Validation result</returns>
        ValidationResult ValidateData(List<Dictionary<string, object>> data);

        /// <summary>
        /// Validates individual record
        /// </summary>
        /// <param name="record">Record to validate</param>
        /// <param name="recordIndex">Index of record for error reporting</param>
        /// <returns>Validation result</returns>
        ValidationResult ValidateRecord(Dictionary<string, object> record, int recordIndex);

        /// <summary>
        /// Validates header consistency
        /// </summary>
        /// <param name="headers">Headers to validate</param>
        /// <returns>Validation result</returns>
        ValidationResult ValidateHeaders(IEnumerable<string> headers);
    }

}
