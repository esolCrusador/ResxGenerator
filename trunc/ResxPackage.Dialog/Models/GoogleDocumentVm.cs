namespace ResxPackage.Dialog.Models
{
    public class GoogleDocumentVm
    {
        public GoogleDocumentVm(string documntPath, string documentUrl, string documentName)
        {
            DocumntPath = documntPath;
            DocumentUrl = documentUrl;
            DocumentName = documentName;
        }

        public string DocumntPath { get; set; }

        public string DocumentUrl { get; set; }

        public string DocumentName { get; set; }

        public bool IsSelected { get; set; }
    }
}
