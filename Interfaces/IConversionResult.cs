using ExportExcel.Models;
using System;
using System.Collections.Generic;


namespace ExportExcel.Interfaces
{
    /// <summary>
    /// Interface for conversion result
    /// </summary>
    public interface IConversionResult
    {
        bool Success { get; }
        string JsonString { get; }
        object Data { get; }
        string ErrorMessage { get; }
        Exception Exception { get; }
        int RecordCount { get; }
        int SheetCount { get; }
        int GroupCount { get; }
        ConversionMode ConversionMode { get; }
        Dictionary<string, object> Metadata { get; }
    }

}
