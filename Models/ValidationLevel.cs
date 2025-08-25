

namespace ExportExcel.Models
{
    /// <summary>
    /// Data validation levels
    /// </summary>
    public enum ValidationLevel
    {
        /// <summary>
        /// No validation
        /// </summary>
        None,

        /// <summary>
        /// Basic validation (null checks, type checks)
        /// </summary>
        Basic,

        /// <summary>
        /// Standard validation (includes data consistency)
        /// </summary>
        Standard,

        /// <summary>
        /// Strict validation (comprehensive checks)
        /// </summary>
        Strict
    }

}
