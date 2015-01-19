using System.Collections.Generic;

namespace Common.Excel.Models
{
    public class ResGroupModel<TModel> where TModel : IRowModel
    {
        public string GroupTitle { get; set; }

        public IReadOnlyList<ResTableModel<TModel>> Tables { get; set; }
    }
}
