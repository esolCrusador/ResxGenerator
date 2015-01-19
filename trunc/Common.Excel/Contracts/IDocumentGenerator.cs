using System.Collections.Generic;
using System.Threading.Tasks;
using Common.Excel.Models;

namespace Common.Excel.Contracts
{
    public interface IDocumentGenerator
    {
        Task ExportToDocumentAsync<TModel>(string path, IReadOnlyList<ResGroupModel<TModel>> groups) where TModel : IRowModel;

        Task<IReadOnlyList<ResGroupModel<TModel>>> ImportFromExcelAsync<TModel>(string path) where TModel : IRowModel, new();
    }
}