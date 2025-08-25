

namespace ExportExcel.Models
{
    /// <summary>
    /// Export modes for different data structures
    /// </summary>
    public enum ExportMode
    {
        /// <summary>
        /// Simple dictionary export
        /// </summary>
        Simple,

        /// <summary>
        /// JSON to Excel conversion
        /// </summary>
        JsonToExcel,

        /// <summary>
        /// Grouped exports based on nested structure
        /// </summary>
        GroupedJson,

        /// <summary>
        /// Flattened JSON exports
        /// </summary>
        FlattenedJson,

        /// <summary>
        /// Combined flattened and grouped
        /// </summary>
        FlattenedGroupedJson
    }

}
