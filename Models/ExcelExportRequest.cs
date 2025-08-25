using System.Collections.Generic;


namespace ExportExcel.Models
{
    public class ExcelExportRequest
    {
        public List<Dictionary<string, object>> Data { get; set; }
        public ExcelExportOptions Options { get; set; } = new ExcelExportOptions();
    }

}
