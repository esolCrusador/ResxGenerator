namespace Common.Excel.Models
{
    public class CellModel : ICellModel
    {
        public bool Hilight { get; set; }

        public string Model { get; set; }

        public string DataString
        {
            get { return Model; }
            set { Model = value; }
        }
    }
}
