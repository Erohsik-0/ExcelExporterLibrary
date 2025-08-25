using System;
using System.Collections.Generic;
using System.Linq;


namespace ExportExcel.Models
{
    /// <summary>
    /// Structure signature information
    /// </summary>
    public class StructureSignature
    {
        public HashSet<string> ObjectProperties { get; set; } = new HashSet<string>();
        public Dictionary<string, int> ArrayProperties { get; set; } = new Dictionary<string, int>();
        public HashSet<string> SimpleProperties { get; set; } = new HashSet<string>();
        public StructureComplexity Complexity { get; set; }
        public int TotalComplexity => ObjectProperties.Count + ArrayProperties.Count;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        public string GetSignatureKey()
        {
            var parts = new List<string>();

            if (ObjectProperties.Count > 0)
            {
                parts.Add($"Objects[{string.Join(",", ObjectProperties.OrderBy(x => x))}]");
            }

            if (ArrayProperties.Count > 0)
            {
                var arrayParts = ArrayProperties.OrderBy(x => x.Key)
                    .Select(x => $"{x.Key}({x.Value})");
                parts.Add($"Arrays[{string.Join(",", arrayParts)}]");
            }

            if (parts.Count == 0)
            {
                return "SimpleRecord";
            }

            return string.Join("_", parts);
        }

        public string GetFriendlyName()
        {
            if (TotalComplexity == 0)
                return "Simple Records";

            var nameParts = new List<string>();

            // Create meaningful names based on common patterns
            if (ObjectProperties.Contains("address") || ObjectProperties.Contains("Address"))
                nameParts.Add("WithAddress");

            if (ObjectProperties.Contains("user") || ObjectProperties.Contains("User") ||
                ObjectProperties.Contains("customer") || ObjectProperties.Contains("Customer"))
                nameParts.Add("WithUser");

            if (ArrayProperties.ContainsKey("orders") || ArrayProperties.ContainsKey("Orders"))
                nameParts.Add("WithOrders");

            if (ArrayProperties.ContainsKey("items") || ArrayProperties.ContainsKey("Items"))
                nameParts.Add("WithItems");

            if (ObjectProperties.Contains("metadata") || ObjectProperties.Contains("Metadata"))
                nameParts.Add("WithMetadata");

            if (nameParts.Count == 0)
            {
                if (ObjectProperties.Count > 0 && ArrayProperties.Count > 0)
                    nameParts.Add("ComplexNested");
                else if (ObjectProperties.Count > 0)
                    nameParts.Add("WithObjects");
                else if (ArrayProperties.Count > 0)
                    nameParts.Add("WithArrays");
            }

            return nameParts.Count > 0 ? string.Join("_", nameParts) : "MixedStructure";
        }

        public double CalculateSimilarity(StructureSignature other)
        {
            if (other == null)
                return 0.0;

            var thisKeys = new HashSet<string>(ObjectProperties);
            thisKeys.UnionWith(ArrayProperties.Keys);
            thisKeys.UnionWith(SimpleProperties);

            var otherKeys = new HashSet<string>(other.ObjectProperties);
            otherKeys.UnionWith(other.ArrayProperties.Keys);
            otherKeys.UnionWith(other.SimpleProperties);

            var intersection = thisKeys.Intersect(otherKeys).Count();
            var union = thisKeys.Union(otherKeys).Count();

            return union > 0 ? (double)intersection / union : 0.0;
        }
    }
}
