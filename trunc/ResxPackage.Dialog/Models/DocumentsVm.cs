using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ResxPackage.Dialog.Models
{
    public class DocumentsVm
    {
        public DocumentsVm(IEnumerable<GoogleDocumentVm>  documents)
        {
            Documents = new ObservableCollection<GoogleDocumentVm>(documents);
        }

        public DocumentsVm()
        {
            Documents = new ObservableCollection<GoogleDocumentVm>();
        }

        public string GetSelectedPath()
        {
            return Documents.Where(doc => doc.IsSelected).Select(doc => doc.DocumntPath).DefaultIfEmpty(null).Single(); 
        }

        public ObservableCollection<GoogleDocumentVm> Documents { get; set; }

        public string GetSelectedUrl()
        {
            return Documents.Where(doc => doc.IsSelected).Select(doc => doc.DocumentUrl).DefaultIfEmpty(null).Single();
        }
    }
}
