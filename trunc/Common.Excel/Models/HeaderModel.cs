using System.Collections.Generic;

namespace Common.Excel.Models
{
    public class HeaderModel<TModel>
    {
        public IReadOnlyList<ColumnModel>  Columns { get; set; }
    }
}
