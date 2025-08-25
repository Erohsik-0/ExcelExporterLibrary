using ExportExcel.Interfaces;
using ExportExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportExcel.Services
{
    /// <summary>
    /// Analyzes data structures for intelligent grouping
    /// </summary>
    public class StructureAnalyzer : IStructureAnalyzer
    {
        private readonly Dictionary<string, StructureSignature> _signatureCache;

        public StructureAnalyzer()
        {
            _signatureCache = new Dictionary<string, StructureSignature>();
        }

        public Dictionary<string, List<Dictionary<string, object>>> GroupByStructure(List<Dictionary<string, object>> data)
        {
            if (data == null || data.Count == 0)
                return new Dictionary<string, List<Dictionary<string, object>>>();

            var groups = new Dictionary<string, List<Dictionary<string, object>>>();

            foreach (var record in data)
            {
                var signature = GetStructureSignature(record);
                var groupKey = GetGroupKeyFromSignature(signature);

                if (!groups.ContainsKey(groupKey))
                {
                    groups[groupKey] = new List<Dictionary<string, object>>();
                }

                groups[groupKey].Add(record);
            }

            // If we only have one group and more than 10 records, try to sub-group
            if (groups.Count == 1 && data.Count > 10)
            {
                var singleGroup = groups.First();
                var subGroups = TryCreateSubGroups(singleGroup.Value);
                if (subGroups.Count > 1)
                {
                    return subGroups;
                }
            }

            return RenameGroupsForClarity(groups);
        }

        public string GetStructureSignature(Dictionary<string, object> record)
        {
            if (record == null || record.Count == 0)
                return "Empty";

            var signature = new StringBuilder();
            var typeCount = new Dictionary<string, int>();
            var complexProperties = new List<string>();

            foreach (var kvp in record.OrderBy(x => x.Key))
            {
                var type = GetValueType(kvp.Value);
                var key = kvp.Key;

                // Track complex nested properties
                if (IsComplexProperty(key, kvp.Value))
                {
                    complexProperties.Add(key);
                }

                if (!typeCount.ContainsKey(type))
                    typeCount[type] = 0;

                typeCount[type]++;
            }

            // Build signature with complex properties first
            if (complexProperties.Count > 0)
            {
                signature.Append($"Complex[{string.Join(",", complexProperties.Take(3))}]");
            }

            foreach (var tc in typeCount.OrderBy(x => x.Key))
            {
                signature.Append($"{tc.Key}:{tc.Value},");
            }

            return signature.ToString().TrimEnd(',');
        }

        public bool AreSimilarStructures(string signature1, string signature2, double similarityThreshold = 0.7)
        {
            if (string.IsNullOrWhiteSpace(signature1) || string.IsNullOrWhiteSpace(signature2))
                return false;

            if (signature1 == signature2)
                return true;

            // Parse signatures and compare
            var sig1Parts = ParseSignature(signature1);
            var sig2Parts = ParseSignature(signature2);

            var intersection = sig1Parts.Intersect(sig2Parts).Count();
            var union = sig1Parts.Union(sig2Parts).Count();

            return union > 0 && (double)intersection / union >= similarityThreshold;
        }

        #region Private Methods

        private string GetValueType(object value)
        {
            if (value == null) return "null";

            var type = value.GetType();

            if (type == typeof(string)) return "string";
            if (type == typeof(int) || type == typeof(long) || type == typeof(decimal) || type == typeof(double)) return "number";
            if (type == typeof(bool)) return "boolean";
            if (type == typeof(DateTime)) return "datetime";
            if (type == typeof(Guid)) return "guid";

            // Handle complex types
            if (value.ToString().StartsWith("{") || value.ToString().StartsWith("["))
                return "json";

            return "object";
        }

        private bool IsComplexProperty(string key, object value)
        {
            // Check if key suggests nested structure
            if (key.Contains(".") || key.Contains("[") || key.Contains("]"))
                return true;

            // Check if value is complex
            if (value == null) return false;

            var valueStr = value.ToString();
            return valueStr.StartsWith("{") || valueStr.StartsWith("[") ||
                   key.ToLower().Contains("address") ||
                   key.ToLower().Contains("user") ||
                   key.ToLower().Contains("order") ||
                   key.ToLower().Contains("metadata");
        }

        private string GetGroupKeyFromSignature(string signature)
        {
            if (string.IsNullOrWhiteSpace(signature))
                return "Unknown";

            var lower = signature.ToLower();

            // Pattern matching for common structures
            if (lower.Contains("complex") && lower.Contains("address"))
                return "WithAddress";
            else if (lower.Contains("complex") && lower.Contains("user"))
                return "WithUser";
            else if (lower.Contains("complex") && lower.Contains("order"))
                return "WithOrders";
            else if (lower.Contains("complex") && lower.Contains("metadata"))
                return "WithMetadata";
            else if (lower.Contains("complex"))
                return "ComplexStructure";
            else if (lower.Contains("json"))
                return "JsonFields";
            else if (lower.Contains("datetime"))
                return "WithDates";
            else if (lower.Contains("number") && !lower.Contains("string"))
                return "NumericData";
            else if (lower.Contains("string") && !lower.Contains("number"))
                return "TextData";

            return "MixedData";
        }

        private Dictionary<string, List<Dictionary<string, object>>> TryCreateSubGroups(List<Dictionary<string, object>> records)
        {
            var subGroups = new Dictionary<string, List<Dictionary<string, object>>>();

            // Try grouping by specific field patterns
            var groupingStrategies = new[]
            {
                // Group by document type or category
                new Func<Dictionary<string, object>, string>(data =>
                    GetValueAsString(data, "type", "documentType", "category", "kind")),
                
                // Group by status or state
                new Func<Dictionary<string, object>, string>(data =>
                    GetValueAsString(data, "status", "state", "condition")),
                
                // Group by presence of specific complex properties
                new Func<Dictionary<string, object>, string>(data =>
                {
                    if (data.Keys.Any(k => k.Contains("order") || k.Contains("Order")))
                        return "WithOrders";
                    if (data.Keys.Any(k => k.Contains("address") || k.Contains("Address")))
                        return "WithAddress";
                    if (data.Keys.Any(k => k.Contains("user") || k.Contains("User")))
                        return "WithUser";
                    return "Standard";
                })
            };

            foreach (var strategy in groupingStrategies)
            {
                var testGroups = new Dictionary<string, List<Dictionary<string, object>>>();

                foreach (var record in records)
                {
                    var groupKey = strategy(record) ?? "Unknown";

                    if (!testGroups.ContainsKey(groupKey))
                        testGroups[groupKey] = new List<Dictionary<string, object>>();

                    testGroups[groupKey].Add(record);
                }

                // If this strategy creates meaningful groups (more than 1 group, no group too dominant)
                if (testGroups.Count > 1 && testGroups.Values.Max(g => g.Count) < records.Count * 0.8)
                {
                    return testGroups;
                }
            }

            // No meaningful sub-grouping found
            return new Dictionary<string, List<Dictionary<string, object>>> { { "All Records", records } };
        }

        private string GetValueAsString(Dictionary<string, object> data, params string[] possibleKeys)
        {
            foreach (var key in possibleKeys)
            {
                if (data.TryGetValue(key, out var value) && value != null)
                {
                    return value.ToString();
                }

                // Try case-insensitive match
                var foundKey = data.Keys.FirstOrDefault(k =>
                    string.Equals(k, key, StringComparison.OrdinalIgnoreCase));

                if (foundKey != null && data[foundKey] != null)
                {
                    return data[foundKey].ToString();
                }
            }

            return null;
        }

        private Dictionary<string, List<Dictionary<string, object>>> RenameGroupsForClarity(
            Dictionary<string, List<Dictionary<string, object>>> groups)
        {
            var renamedGroups = new Dictionary<string, List<Dictionary<string, object>>>();
            int groupIndex = 1;

            foreach (var group in groups.OrderByDescending(g => g.Value.Count))
            {
                var baseName = group.Key;
                var count = group.Value.Count;

                // Create descriptive names
                var name = count > 1 ?
                    $"{baseName}_({count}_Records)" :
                    $"{baseName}_SingleRecord";

                renamedGroups[name] = group.Value;
                groupIndex++;
            }

            return renamedGroups;
        }

        private HashSet<string> ParseSignature(string signature)
        {
            var parts = new HashSet<string>();

            if (string.IsNullOrWhiteSpace(signature))
                return parts;

            var segments = signature.Split(',');
            foreach (var segment in segments)
            {
                var normalized = segment.Trim().ToLower();
                if (!string.IsNullOrEmpty(normalized))
                {
                    parts.Add(normalized);
                }
            }

            return parts;
        }

        #endregion
    }
}
