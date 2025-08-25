using System.Collections.Generic;


namespace ExportExcel.Interfaces
{
    /// <summary>
    /// Interface for validation result
    /// </summary>
    public interface IValidationResult
    {
        bool IsValid { get; }
        List<string> Errors { get; }
        List<string> Warnings { get; }
        Dictionary<string, object> ValidationMetadata { get; }
    }

}
