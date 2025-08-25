using System;
using System.Collections.Generic;


namespace ExportExcel.Models
{
    /// <summary>
    /// Represents a record with structure information
    /// </summary>
    public class StructuredRecord
    {
        public Dictionary<string, object> FlatData { get; set; }
        public StructureSignature Structure { get; set; }
        public object OriginalToken { get; set; }
        public int RecordIndex { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        public StructuredRecord()
        {
            FlatData = new Dictionary<string, object>();
            Structure = new StructureSignature();
        }
    }

}
