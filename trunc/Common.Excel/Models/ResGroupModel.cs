using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common.Excel.Models
{
    public class ResGroupModel<TModel> where TModel : IRowModel
    {
        public string GroupTitle { get; set; }

        public IReadOnlyList<ResTableModel<TModel>> Tables { get; set; }
    }
}
