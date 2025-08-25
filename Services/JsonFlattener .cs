using ExportExcel.Exceptions;
using ExportExcel.Interfaces;
using ExportExcel.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExportExcel.Services
{
    /// <summary>
    /// JSON flattening and structure reconstruction service
    /// </summary>
    public class JsonFlattener : IJsonFlattener
    {
        private readonly ExcelImportOptions _options;

        public JsonFlattener(ExcelImportOptions options = null)
        {
            _options = options ?? new ExcelImportOptions();
        }

        public List<Dictionary<string, object>> FlattenJson(string jsonString)
        {
            if (string.IsNullOrWhiteSpace(jsonString))
                return new List<Dictionary<string, object>>();

            try
            {
                var flatList = new List<Dictionary<string, object>>();
                var orderedHeaders = new List<string>();
                bool isFirstItem = true;

                var token = JToken.Parse(jsonString);

                if (token is JObject obj)
                {
                    // Look for common array containers
                    var arrayContainers = new[] { "Documents", "Items", "Data", "Records", "Results", "Entities" };
                    JArray targetArray = null;

                    foreach (var container in arrayContainers)
                    {
                        if (obj[container] is JArray arr)
                        {
                            targetArray = arr;
                            break;
                        }
                    }

                    if (targetArray != null)
                    {
                        ProcessJsonArray(targetArray, flatList, ref orderedHeaders, ref isFirstItem);
                    }
                    else
                    {
                        // Single object case
                        flatList.Add(FlattenJsonObject(obj, "", isFirstItem, ref orderedHeaders));
                    }
                }
                else if (token is JArray arr)
                {
                    ProcessJsonArray(arr, flatList, ref orderedHeaders, ref isFirstItem);
                }

                return flatList;
            }
            catch (JsonReaderException ex)
            {
                throw new JsonParsingException(
                    $"Invalid JSON format: {ex.Message}",
                    jsonString?.Substring(0, Math.Min(jsonString.Length, 100)),
                    ex,
                    ex.LineNumber,
                    ex.LinePosition);
            }
        }

        public List<Dictionary<string, object>> ReconstructNestedStructure(List<Dictionary<string, object>> flatData)
        {
            if (flatData == null || flatData.Count == 0)
                return new List<Dictionary<string, object>>();

            var result = new List<Dictionary<string, object>>();

            foreach (var flatRecord in flatData)
            {
                var nestedRecord = new Dictionary<string, object>();

                foreach (var kvp in flatRecord)
                {
                    SetNestedValue(nestedRecord, kvp.Key, kvp.Value);
                }

                result.Add(nestedRecord);
            }

            return result;
        }

        public bool IsNestedField(string fieldName)
        {
            if (string.IsNullOrWhiteSpace(fieldName))
                return false;

            // Check for dot notation or array notation
            return fieldName.Contains(".") || fieldName.Contains("[") || fieldName.Contains("]");
        }

        

        private void ProcessJsonArray(JArray array, List<Dictionary<string, object>> flatList,
            ref List<string> orderedHeaders, ref bool isFirstItem)
        {
            foreach (var item in array)
            {
                flatList.Add(FlattenJsonObject(item, "", isFirstItem, ref orderedHeaders));
                isFirstItem = false;
            }
        }

        private Dictionary<string, object> FlattenJsonObject(JToken token, string parentPath, bool isFirstItem, ref List<string> orderedHeaders)
        {
            var result = new Dictionary<string, object>();

            if (token is JObject jObj)
            {
                foreach (var prop in jObj.Properties())
                {
                    string path = string.IsNullOrEmpty(parentPath) ? prop.Name : $"{parentPath}.{prop.Name}";
                    var value = prop.Value;

                    switch (value.Type)
                    {
                        case JTokenType.Object:
                            // Recursively flatten nested objects
                            var nestedResult = FlattenJsonObject(value, path, isFirstItem, ref orderedHeaders);
                            foreach (var nested in nestedResult)
                            {
                                result[nested.Key] = nested.Value;
                                if (isFirstItem && !orderedHeaders.Contains(nested.Key))
                                    orderedHeaders.Add(nested.Key);
                            }
                            break;

                        case JTokenType.Array:
                            // Handle arrays by creating indexed keys
                            var array = value as JArray;
                            ProcessArrayToken(array, path, result, isFirstItem, ref orderedHeaders);
                            break;

                        case JTokenType.Null:
                            result[path] = null;
                            if (isFirstItem && !orderedHeaders.Contains(path))
                                orderedHeaders.Add(path);
                            break;

                        default:
                            // Simple values - just add them
                            result[path] = ((JValue)value).Value;
                            if (isFirstItem && !orderedHeaders.Contains(path))
                                orderedHeaders.Add(path);
                            break;
                    }
                }
            }
            else if (token is JArray directArray)
            {
                ProcessArrayToken(directArray, parentPath, result, isFirstItem, ref orderedHeaders);
            }

            return result;
        }

        private void ProcessArrayToken(JArray array, string path, Dictionary<string, object> result,
            bool isFirstItem, ref List<string> orderedHeaders)
        {
            // For small arrays, expand each element
            if (array.Count <= 10)
            {
                for (int i = 0; i < array.Count; i++)
                {
                    if (array[i] is JObject itemObj)
                    {
                        var arrayResult = FlattenJsonObject(itemObj, $"{path}[{i}]", isFirstItem, ref orderedHeaders);
                        foreach (var nested in arrayResult)
                        {
                            result[nested.Key] = nested.Value;
                            if (isFirstItem && !orderedHeaders.Contains(nested.Key))
                                orderedHeaders.Add(nested.Key);
                        }
                    }
                    else
                    {
                        string arrayKey = $"{path}[{i}]";
                        result[arrayKey] = ((JValue)array[i])?.Value;
                        if (isFirstItem && !orderedHeaders.Contains(arrayKey))
                            orderedHeaders.Add(arrayKey);
                    }
                }
            }
            else
            {
                // For large arrays, create summary information
                result[$"{path}_Count"] = array.Count;
                if (isFirstItem && !orderedHeaders.Contains($"{path}_Count"))
                    orderedHeaders.Add($"{path}_Count");

                // Sample first few elements
                for (int i = 0; i < Math.Min(3, array.Count); i++)
                {
                    if (array[i] is JObject itemObj)
                    {
                        var arrayResult = FlattenJsonObject(itemObj, $"{path}[{i}]", isFirstItem, ref orderedHeaders);
                        foreach (var nested in arrayResult)
                        {
                            result[nested.Key] = nested.Value;
                            if (isFirstItem && !orderedHeaders.Contains(nested.Key))
                                orderedHeaders.Add(nested.Key);
                        }
                    }
                    else
                    {
                        string arrayKey = $"{path}[{i}]";
                        result[arrayKey] = ((JValue)array[i])?.Value;
                        if (isFirstItem && !orderedHeaders.Contains(arrayKey))
                            orderedHeaders.Add(arrayKey);
                    }
                }
            }
        }

        private void SetNestedValue(Dictionary<string, object> target, string path, object value)
        {
            var parts = ParsePath(path);
            var current = target;

            for (int i = 0; i < parts.Count - 1; i++)
            {
                var part = parts[i];

                if (part.IsArray)
                {
                    // Handle array notation
                    if (!current.ContainsKey(part.Name))
                    {
                        current[part.Name] = new List<Dictionary<string, object>>();
                    }

                    var list = current[part.Name] as List<Dictionary<string, object>>;

                    // Ensure list has enough elements
                    while (list.Count <= part.Index)
                    {
                        list.Add(new Dictionary<string, object>());
                    }

                    current = list[part.Index];
                }
                else
                {
                    // Handle object notation
                    if (!current.ContainsKey(part.Name))
                    {
                        current[part.Name] = new Dictionary<string, object>();
                    }

                    current = current[part.Name] as Dictionary<string, object>;
                }
            }

            // Set the final value
            var lastPart = parts.Last();
            if (lastPart.IsArray)
            {
                if (!current.ContainsKey(lastPart.Name))
                {
                    current[lastPart.Name] = new List<object>();
                }

                var list = current[lastPart.Name] as List<object>;
                while (list.Count <= lastPart.Index)
                {
                    list.Add(null);
                }

                list[lastPart.Index] = value;
            }
            else
            {
                current[lastPart.Name] = value;
            }
        }

        private List<PathPart> ParsePath(string path)
        {
            var parts = new List<PathPart>();
            var regex = new Regex(@"([^\.\[]+)(\[(\d+)\])?");
            var segments = path.Split('.');

            foreach (var segment in segments)
            {
                var matches = regex.Matches(segment);
                foreach (Match match in matches)
                {
                    var part = new PathPart
                    {
                        Name = match.Groups[1].Value,
                        FullPath = path
                    };

                    if (match.Groups[3].Success)
                    {
                        part.IsArray = true;
                        part.Index = int.Parse(match.Groups[3].Value);
                    }

                    parts.Add(part);
                }
            }

            return parts;
        }

    }
}
