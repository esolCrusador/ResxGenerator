using System.Collections.Generic;

namespace Common.Excel.Models
{
    public abstract class RowModel : IRowModel
    {
        public string Color { get; set; }

        public abstract IReadOnlyList<ICellModel> DataList { get; set; }
    }

    public class RowModel<TModel> : RowModel
        where TModel : IRowModel
    {
        public TModel Model { get; set; }

        public override IReadOnlyList<ICellModel> DataList
        {
            get { return Model.DataList; }
            set { Model.DataList = value; }
        }
    }
}
