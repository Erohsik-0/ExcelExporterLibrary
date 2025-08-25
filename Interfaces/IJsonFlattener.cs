using System.Collections.Generic;


namespace ExportExcel.Interfaces
{
    /// <summary>
    /// Interface for JSON flattening operations
    /// </summary>
    public interface IJsonFlattener
    {
        /// <summary>
        /// Flattens nested JSON objects into dot-notation dictionary
        /// </summary>
        /// <param name="jsonString">JSON string to flatten</param>
        /// <returns>List of flattened dictionaries</returns>
        List<Dictionary<string, object>> FlattenJson(string jsonString);

        /// <summary>
        /// Reconstructs nested structure from flattened data
        /// </summary>
        /// <param name="flatData">Flattened data</param>
        /// <returns>List of nested dictionaries</returns>
        List<Dictionary<string, object>> ReconstructNestedStructure(List<Dictionary<string, object>> flatData);

        /// <summary>
        /// Determines if a field name represents nested data
        /// </summary>
        /// <param name="fieldName">Field name to check</param>
        /// <returns>True if field represents nested data</returns>
        bool IsNestedField(string fieldName);
    }

}
