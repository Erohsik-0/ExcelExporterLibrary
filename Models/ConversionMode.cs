

namespace ExportExcel.Models
{
    /// <summary>
    /// Conversion modes for different Excel structures
    /// </summary>
    public enum ConversionMode
    {
        /// <summary>
        /// Automatically detect best mode
        /// </summary>
        Auto,

        /// <summary>
        /// Simple flat structure
        /// </summary>
        Simple,

        /// <summary>
        /// Reconstruct nested objects from flat columns
        /// </summary>
        Nested,

        /// <summary>
        /// Multiple sheets as separate arrays
        /// </summary>
        MultiSheet,

        /// <summary>
        /// Group by structure similarity
        /// </summary>
        Grouped
    }

}
