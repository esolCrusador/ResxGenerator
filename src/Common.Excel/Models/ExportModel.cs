using System.Collections.Generic;

namespace Common.Excel.Models
{
    public class ExportModel
    {
        public string Title { get; set; }
        public List<string> ColumnHeaders { get; set; }
        public List<List<string>> Rows{ get; set; }

        public ExportModel()
        {
            Rows = new List<List<string>>();
            ColumnHeaders = new List<string>();
            Title = "Exported Data";
        }
    }
}
