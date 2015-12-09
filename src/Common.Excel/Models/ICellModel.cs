namespace Common.Excel.Models
{
    public interface ICellModel
    {
        bool Hilight { get; set; }

        string DataString { get; set; }
    }
}
