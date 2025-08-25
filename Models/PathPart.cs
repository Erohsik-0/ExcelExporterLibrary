

namespace ExportExcel.Models
{
    /// <summary>
    /// Path parsing information for nested structures
    /// </summary>
    public class PathPart
    {
        public string Name { get; set; }
        public bool IsArray { get; set; }
        public int Index { get; set; }
        public string FullPath { get; set; }

        public override string ToString()
        {
            return IsArray ? $"{Name}[{Index}]" : Name;
        }
    }

}
