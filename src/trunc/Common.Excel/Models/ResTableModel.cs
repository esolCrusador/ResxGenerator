using System.Collections.Generic;

namespace Common.Excel.Models
{
    public class ResTableModel<TModel> where TModel : IRowModel
    {
        public string TableTitle { get; set; }

        public HeaderModel Header { get; set; }

        public IReadOnlyList<RowModel<TModel>> Rows { get; set; }
    }
}
