namespace Common.Excel.Models
{
    public class DocumentInfoModel
    {
        public DocumentInfoModel(string path, string url, string name)
        {
            Path = path;
            Url = url;
            Name = name;
        }

        public string Path { get; set; }

        public string Url { get; set; }

        public string Name { get; set; }
    }
}
