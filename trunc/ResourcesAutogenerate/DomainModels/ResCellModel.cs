using Common.Excel.Models;

namespace ResourcesAutogenerate.DomainModels
{
    public class ResCellModel : ICellModel
    {
        public ResCellModel(string resourceKey)
        {
            ResValue = resourceKey;
        }

        public string ResValue { get; set; }

        public bool Hilight { get; set; }

        public string DataString
        {
            get { return ResValue; }
            set { ResValue = value; }
        }
    }
}
