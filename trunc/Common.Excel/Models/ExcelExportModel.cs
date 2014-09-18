using System.Collections.Generic;

namespace Common.Excel.Models
{
    public class ExcelExportModel
    {
        public string Title { get; set; }
        public List<string> ColumnHeaders { get; set; }
        public List<List<string>> Rows{ get; set; }

        public ExcelExportModel()
        {
            Rows = new List<List<string>>();
            ColumnHeaders = new List<string>();
            Title = "Exported Data";
        }
    }
}
