using System.Collections.Generic;
using ExportExcel.Interfaces;

namespace ExportExcel.Models
{
    /// <summary>
    /// Result of data validation operations
    /// </summary>
    public class ValidationResult : IValidationResult
    {
        public bool IsValid { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();
        public Dictionary<string, object> ValidationMetadata { get; set; } = new Dictionary<string, object>();

        public ValidationResult()
        {
            IsValid = true;
        }

        public ValidationResult(bool isValid)
        {
            IsValid = isValid;
        }

        public void AddError(string error)
        {
            Errors.Add(error);
            IsValid = false;
        }

        public void AddWarning(string warning)
        {
            Warnings.Add(warning);
        }

        public void AddMetadata(string key, object value)
        {
            ValidationMetadata[key] = value;
        }

    }

}
