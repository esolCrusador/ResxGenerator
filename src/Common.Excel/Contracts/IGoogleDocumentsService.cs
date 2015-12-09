using System.Collections.Generic;
using System.Threading.Tasks;
using Common.Excel.Models;

namespace Common.Excel.Contracts
{
    public interface IGoogleDocumentsService
    {
        Task<IReadOnlyCollection<DocumentInfoModel>> GetDocuments();
    }
}
