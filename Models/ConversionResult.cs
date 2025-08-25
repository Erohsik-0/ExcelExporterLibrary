using System;
using System.Collections.Generic;
using ExportExcel.Interfaces;

namespace ExportExcel.Models
{
    /// <summary>
    /// Result of Excel conversion operations
    /// </summary>
    public class ConversionResult : IConversionResult
    {
        public bool Success { get; set; }
        public string JsonString { get; set; }
        public object Data { get; set; }
        public string ErrorMessage { get; set; }
        public Exception Exception { get; set; }
        public int RecordCount { get; set; }
        public int SheetCount { get; set; }
        public int GroupCount { get; set; }
        public ConversionMode ConversionMode { get; set; }
        public Dictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();

        public ConversionResult()
        {
            Success = false;
            RecordCount = 0;
            SheetCount = 0;
            GroupCount = 0;
            ConversionMode = ConversionMode.Auto;
        }

        public void AddMetadata(string key, object value)
        {
            Metadata[key] = value;
        }

        public T GetMetadata<T>(string key, T defaultValue = default(T))
        {
            if (Metadata.TryGetValue(key, out var value) && value is T)
            {
                return (T)value;
            }
            return defaultValue;
        }
    }
}
