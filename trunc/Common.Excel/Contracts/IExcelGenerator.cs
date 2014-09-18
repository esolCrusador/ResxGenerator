using System.Collections.Generic;
using Common.Excel.Models;

namespace Common.Excel.Contracts
{
    public interface IExcelGenerator
    {
        FileInfoContainer ExportToExcel(ExcelExportModel mdl);

        FileInfoContainer ExportToExcel<TModel>(IReadOnlyList<ResGroupModel<TModel>> groups, string title) where TModel : IRowModel;

        IReadOnlyList<ResGroupModel<TModel>> ImportFromExcel<TModel>(FileInfoContainer file) where TModel : IRowModel, new();
    }
}