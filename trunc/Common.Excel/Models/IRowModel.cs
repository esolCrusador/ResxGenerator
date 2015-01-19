using System.Collections.Generic;

namespace Common.Excel.Models
{
    public interface IRowModel
    {
        IReadOnlyList<ICellModel> DataList { get; set; }
    }
}
