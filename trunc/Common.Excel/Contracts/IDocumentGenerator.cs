using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Common.Excel.Models;

namespace Common.Excel.Contracts
{
    public interface IDocumentGenerator
    {
        Task ExportToDocumentAsync<TModel>(string path, IReadOnlyList<ResGroupModel<TModel>> groups, IStatusProgress progress, CancellationToken cancellationToken) where TModel : IRowModel;

        Task<IReadOnlyList<ResGroupModel<TModel>>> ImportFromExcelAsync<TModel>(string path, IStatusProgress progress, CancellationToken cancellationToken) where TModel : IRowModel, new();
    }
}