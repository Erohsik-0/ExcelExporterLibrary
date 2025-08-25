using ClosedXML.Excel;
using ExportExcel.Interfaces;
using ExportExcel.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExportExcel.Services
{
    /// <summary>
    /// Intelligent type detection and conversion for Excel cells
    /// </summary>
    public class TypeDetector : ITypeDetector
    {
        private readonly ExcelImportOptions _options;
        private readonly Dictionary<string, Type> _typeCache;
        private readonly Dictionary<string, string> _dateFormats;

        public TypeDetector(ExcelImportOptions options = null)
        {
            _options = options ?? new ExcelImportOptions();
            _typeCache = new Dictionary<string, Type>();
            _dateFormats = new Dictionary<string, string>
            {
                { "yyyy-mm-dd", "yyyy-MM-dd" },
                { "dd/mm/yyyy", "dd/MM/yyyy" },
                { "mm/dd/yyyy", "MM/dd/yyyy" },
                { "yyyy-mm-dd hh:mm:ss", "yyyy-MM-dd HH:mm:ss" }
            };
        }

        public object DetectAndConvert(IXLCell cell)
        {
            if (cell == null || cell.IsEmpty())
                return null;

            // Try to get the value with Excel's type detection first
            var dataType = cell.DataType;

            switch (dataType)
            {
                case XLDataType.Boolean:
                    return cell.GetValue<bool>();

                case XLDataType.Number:
                    return ConvertNumber(cell);

                case XLDataType.DateTime:
                    return ConvertDateTime(cell);

                case XLDataType.Text:
                    return ConvertText(cell);

                case XLDataType.Blank:
                    return _options.PreserveNullValues ? null : "";

                default:
                    return cell.GetValue<string>();
            }
        }

        public object TryParseSpecialTypes(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return _options.PreserveNullValues ? null : "";

            // Check cache first
            if (_typeCache.TryGetValue(text, out var cachedType))
            {
                return ConvertCachedType(text, cachedType);
            }

            // Try parsing as various types
            var converted = AttemptTypeParsing(text);
            if (converted != null)
            {
                _typeCache[text] = converted.GetType();
                return converted;
            }

            return text;
        }

        public void ClearCache()
        {
            _typeCache.Clear();
        }

        #region Private Methods

        private object ConvertNumber(IXLCell cell)
        {
            var value = cell.GetValue<double>();

            // Check if it's actually a date (Excel stores dates as numbers)
            if (cell.Style.DateFormat != null && !string.IsNullOrEmpty(cell.Style.DateFormat.Format))
            {
                try
                {
                    return DateTime.FromOADate(value);
                }
                catch
                {
                    // If date conversion fails, treat as number
                }
            }

            // Determine if it's an integer or decimal
            if (Math.Abs(value % 1) < double.Epsilon)
            {
                // It's a whole number
                if (value >= int.MinValue && value <= int.MaxValue)
                {
                    return (int)value;
                }
                else if (value >= long.MinValue && value <= long.MaxValue)
                {
                    return (long)value;
                }
            }

            // Return as decimal for precision
            if (_options.UseDecimalForNumbers)
            {
                return (decimal)value;
            }

            return value;
        }

        private object ConvertDateTime(IXLCell cell)
        {
            try
            {
                var dateTime = cell.GetValue<DateTime>();

                if (_options.ConvertDatesToStrings)
                {
                    return dateTime.ToString(_options.DateFormat);
                }

                return dateTime;
            }
            catch
            {
                // If conversion fails, return as string
                return cell.GetValue<string>();
            }
        }

        private object ConvertText(IXLCell cell)
        {
            var text = cell.GetValue<string>();

            if (_options.TrimWhitespace)
            {
                text = text?.Trim();
            }

            if (string.IsNullOrWhiteSpace(text))
                return _options.PreserveNullValues ? null : "";

            // Try to parse special types if configured
            if (_options.AutoDetectTypes)
            {
                return TryParseSpecialTypes(text);
            }

            return text;
        }

        private object AttemptTypeParsing(string text)
        {
            // Boolean
            if (bool.TryParse(text, out var boolValue))
                return boolValue;

            // Integer
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var intValue))
                return intValue;

            // Long
            if (long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var longValue))
                return longValue;

            // Decimal
            if (decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out var decimalValue))
                return _options.UseDecimalForNumbers ? decimalValue : (double)decimalValue;

            // DateTime
            if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateValue))
            {
                return _options.ConvertDatesToStrings ? text : dateValue;
            }

            // GUID
            if (Guid.TryParse(text, out var guidValue))
                return _options.PreserveGuidsAsStrings ? text : guidValue;

            // JSON
            if (_options.ParseJsonStrings && (text.StartsWith("{") || text.StartsWith("[")))
            {
                try
                {
                    return JToken.Parse(text);
                }
                catch
                {
                    // Not valid JSON, continue
                }
            }

            return null;
        }

        private object ConvertCachedType(string text, Type type)
        {
            try
            {
                if (type == typeof(bool))
                    return bool.Parse(text);
                if (type == typeof(int))
                    return int.Parse(text, CultureInfo.InvariantCulture);
                if (type == typeof(long))
                    return long.Parse(text, CultureInfo.InvariantCulture);
                if (type == typeof(decimal))
                    return decimal.Parse(text, CultureInfo.InvariantCulture);
                if (type == typeof(double))
                    return double.Parse(text, CultureInfo.InvariantCulture);
                if (type == typeof(DateTime))
                    return DateTime.Parse(text, CultureInfo.InvariantCulture);
                if (type == typeof(Guid))
                    return _options.PreserveGuidsAsStrings ? text : Guid.Parse(text);
            }
            catch
            {
                // Conversion failed, remove from cache
                _typeCache.Remove(text);
            }

            return text;
        }

        #endregion
    }
}
