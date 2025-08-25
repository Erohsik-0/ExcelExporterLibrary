using ClosedXML.Excel;

namespace ExportExcel.Interfaces
{
    /// <summary>
    /// Interface for type detection and conversion
    /// </summary>
    public interface ITypeDetector
    {
        /// <summary>
        /// Detects and converts Excel cell value to appropriate .NET type
        /// </summary>
        /// <param name="cell">Excel cell to analyze</param>
        /// <returns>Converted value with appropriate type</returns>
        object DetectAndConvert(IXLCell cell);

        /// <summary>
        /// Tries to parse text as various types
        /// </summary>
        /// <param name="text">Text to parse</param>
        /// <returns>Parsed value or null if no type match found</returns>
        object TryParseSpecialTypes(string text);

        /// <summary>
        /// Clears the internal type cache
        /// </summary>
        void ClearCache();
    }

}
