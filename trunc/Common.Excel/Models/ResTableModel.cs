using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common.Excel.Models
{
    public class ResTableModel<TModel> where TModel : IRowModel
    {
        public string TableTitle { get; set; }

        public HeaderModel<TModel> Header { get; set; }

        public IReadOnlyList<RowModel<TModel>> Rows { get; set; }
    }
}
