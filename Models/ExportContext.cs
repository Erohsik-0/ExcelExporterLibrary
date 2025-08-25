using System;
using System.Collections.Generic;


namespace ExportExcel.Models
{
    /// <summary>
    /// Export operation context
    /// </summary>
    public class ExportContext
    {
        public ExportMode Mode { get; set; }
        public ExcelExportOptions Options { get; set; }
        public Dictionary<string, object> Parameters { get; set; } = new Dictionary<string, object>();
        public DateTime StartTime { get; set; } = DateTime.UtcNow;
        public TimeSpan? Duration { get; set; }
        public List<string> ProcessingLog { get; set; } = new List<string>();

        public void AddLogEntry(string message)
        {
            ProcessingLog.Add($"{DateTime.UtcNow:HH:mm:ss.fff}: {message}");
        }

        public void SetDuration()
        {
            Duration = DateTime.UtcNow - StartTime;
        }
    }

}
