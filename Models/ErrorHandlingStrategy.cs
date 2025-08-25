

namespace ExportExcel.Models
{
    /// <summary>
    /// Error handling strategies
    /// </summary>
    public enum ErrorHandlingStrategy
    {
        /// <summary>
        /// Throw exceptions immediately
        /// </summary>
        ThrowImmediately,

        /// <summary>
        /// Collect errors and throw at end
        /// </summary>
        CollectAndThrow,

        /// <summary>
        /// Skip invalid data and continue
        /// </summary>
        SkipInvalidData,

        /// <summary>
        /// Use default values for invalid data
        /// </summary>
        UseDefaultValues
    }

}
