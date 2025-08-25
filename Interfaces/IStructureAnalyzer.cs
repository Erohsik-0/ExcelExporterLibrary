using System.Collections.Generic;

namespace ExportExcel.Interfaces
{
    /// <summary>
    /// Interface for data structure analysis
    /// </summary>
    public interface IStructureAnalyzer
    {
        /// <summary>
        /// Groups records by their structural similarity
        /// </summary>
        /// <param name="data">Data to analyze and group</param>
        /// <returns>Dictionary of groups with structure-based keys</returns>
        Dictionary<string, List<Dictionary<string, object>>> GroupByStructure(List<Dictionary<string, object>> data);

        /// <summary>
        /// Analyzes the structure signature of a single record
        /// </summary>
        /// <param name="record">Record to analyze</param>
        /// <returns>Structure signature string</returns>
        string GetStructureSignature(Dictionary<string, object> record);

        /// <summary>
        /// Determines if two structure signatures are similar
        /// </summary>
        /// <param name="signature1">First signature</param>
        /// <param name="signature2">Second signature</param>
        /// <param name="similarityThreshold">Similarity threshold (0.0 to 1.0)</param>
        /// <returns>True if signatures are similar</returns>
        bool AreSimilarStructures(string signature1, string signature2, double similarityThreshold = 0.7);
    }

}
