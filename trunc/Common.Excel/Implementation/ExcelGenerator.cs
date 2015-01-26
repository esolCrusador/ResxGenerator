using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Common.Excel.Contracts;
using Common.Excel.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ResxPackage.Resources;

namespace Common.Excel.Implementation
{
    public class ExcelGenerator : IDocumentGenerator
    {
        //For Excel2007 and above .xlsx files

        public Task ExportToDocumentAsync<TModel>(string path, IReadOnlyList<ResGroupModel<TModel>> groups, IStatusProgress progress, CancellationToken cancellationToken) where TModel : IRowModel
        {
            return Task.Run(()=> ExportToDocument(path, groups, progress, cancellationToken), cancellationToken);
        }

        public Task<IReadOnlyList<ResGroupModel<TModel>>> ImportFromDocumentAsync<TModel>(string path, IStatusProgress progress, CancellationToken cancellationToken) where TModel : IRowModel, new()
        {
            return Task.Run(() => ImportFromExcel<TModel>(path));
        }

        private void ExportToDocument<TModel>(string path, IReadOnlyList<ResGroupModel<TModel>> groups, IStatusProgress progress, CancellationToken cancellationToken) where TModel : IRowModel
        {
            if (groups.Count == 0)
            {
                throw new ArgumentException("There is not resource files to export", "groups");
            }

            progress.Report(StatusRes.ExportingToExcel);

            using (var stream = new MemoryStream())
            {
                /* Create the worksheet. */
                SpreadsheetDocument spreadsheet = Excel.CreateWorkbook(stream);
                Excel.AddBasicStyles(spreadsheet);
                Excel.AddAdditionalStyles(spreadsheet);

                double totalRows = groups.Sum(g => g.Tables.Sum(t => t.Rows.Count));
                int rowIndexReport = 1;

                for (int projectIndex = 0; projectIndex < groups.Count; projectIndex++)
                {
                    var @group = groups[projectIndex];
                    Excel.AddWorksheet(spreadsheet, @group.GroupTitle);
                    Worksheet worksheet = spreadsheet.WorkbookPart.WorksheetParts.ElementAt(projectIndex).Worksheet;


                    uint rowIndex = 1;

                    int columnsCount = groups.SelectMany(g => g.Tables.Select(t => t.Header.Columns.Count)).Max();
                    List<int> columnWidthes = new List<int>(columnsCount);

                    for (int i = 0; i < columnsCount; i++)
                    {
                        List<string> groupStrings = groups.SelectMany(g =>
                            g.Tables.Select(t =>
                            {
                                string colTitle = t.Header.Columns[i].Title;
                                int longestRowLength = t.Rows.Select(r => r.DataList[i].DataString).Max(str => str.Length);
                                string longestRow = t.Rows.Select(r => r.DataList[i].DataString).First(str => str.Length == longestRowLength);

                                return colTitle.Length > longestRow.Length ? colTitle : longestRow;
                            })
                            )
                            .ToList();

                        int longestStringLength = groupStrings.Max(s => s.Length);
                        string longestString = groupStrings.Find(str => str.Length == longestStringLength);

                        columnWidthes.Add((int) Excel.GetDefaultFontWidth(longestString));
                    }

                    //Setting Columns
                    foreach (var resTableModel in @group.Tables)
                    {
                        Excel.SetColumnHeadingValue(spreadsheet, worksheet, 1, rowIndex++, resTableModel.TableTitle, false, false);
                        Excel.SetCellValue(spreadsheet, worksheet, 1, rowIndex++, " ", false, false);

                        for (int columnIndex = 0; columnIndex < resTableModel.Header.Columns.Count; columnIndex++)
                        {
                            Excel.SetColumnHeadingValue(spreadsheet, worksheet, (uint) columnIndex + 1, rowIndex, resTableModel.Header.Columns[columnIndex].Title, false, false);

                            var width = columnWidthes[columnIndex];
                            if (width > 0)
                            {
                                Excel.SetColumnWidth(worksheet, columnIndex + 1, width);
                            }
                        }

                        rowIndex++;

                        foreach (var tableRow in resTableModel.Rows)
                        {
                            for (int columnIndex = 0; columnIndex < tableRow.DataList.Count; columnIndex++)
                            {
                                var cell = tableRow.DataList[columnIndex];

                                Excel.SetCellValue(spreadsheet, worksheet, (uint) columnIndex + 1, rowIndex, cell.DataString, false, false, highlight: cell.Hilight);
                            }

                            rowIndex++;

                            rowIndexReport++;
                        }

                        progress.Report(100 * rowIndexReport / totalRows);
                        cancellationToken.ThrowIfCancellationRequested();

                        Excel.SetCellValue(spreadsheet, worksheet, 1, rowIndex++, " ", false, false);
                    }

                    worksheet.Save();
                }

                spreadsheet.Close();

                if (File.Exists(path))
                {
                    File.Delete(path);
                }

                using (FileStream fileStream = File.Create(path))
                {
                    fileStream.Write(stream.ToArray(), 0, (int)stream.Length);
                }
            }
        }

        public IReadOnlyList<ResGroupModel<TModel>> ImportFromExcel<TModel>(string path) where TModel : IRowModel, new()
        {
            using (FileStream stream =  File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, false);
                WorkbookPart workbookPart = spreadsheet.WorkbookPart;

                var dataList = new List<ResGroupModel<TModel>>();

                for (int projectIndex = 0; projectIndex < workbookPart.WorksheetParts.Count(); projectIndex++)
                {
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.ElementAt(projectIndex);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    var itemGetter = GetSharedStringItemById(workbookPart);

                    List<List<string>> data = sheetData.Cast<Row>()
                        .Select(row => row.ChildElements.Cast<Cell>()
                            .Select(cell => cell.CellValue != null ? itemGetter(cell.CellValue.Text) : null)
                            .ToList()
                        )
                        .ToList();

                    var dataEnumerator = data.GetEnumerator();
                    List<string> rowCells = null;

                    Func<bool> moveNextFunc = () =>
                        dataEnumerator.MoveNext() && (rowCells = dataEnumerator.Current) != null && rowCells.Count != 0 && !String.IsNullOrWhiteSpace(rowCells[0]);


                    string worpbookPartId = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart);
                    Sheet sheet = workbookPart.Workbook.Sheets.Cast<Sheet>().First(s => s.Id == worpbookPartId);
                    var newGroup = new ResGroupModel<TModel> { GroupTitle = sheet.Name };

                    var tables = new List<ResTableModel<TModel>>();


                    while (moveNextFunc())
                    {
                        var newTable = new ResTableModel<TModel>
                        {
                            TableTitle = rowCells[0]
                        };

                        dataEnumerator.MoveNext();
                        moveNextFunc();

                        newTable.Header = new HeaderModel
                        {
                            Columns = rowCells.Select(cell => new ColumnModel {Title = cell}).ToList()
                        };

                        var rows = new List<RowModel<TModel>>();
                        while (moveNextFunc())
                        {
                            var tableRow = new RowModel<TModel>
                            {
                                Model = new TModel
                                {
                                    DataList = rowCells.Select(cell => new CellModel {Model = cell}).ToList()
                                }
                            };

                            rows.Add(tableRow);
                        }

                        newTable.Rows = rows;

                        tables.Add(newTable);

                        //Space between resources
                    }

                    newGroup.Tables = tables;
                    dataList.Add(newGroup);
                }

                return dataList;
            }
        }

        private static Func<string, string> GetSharedStringItemById(WorkbookPart workbookPart)
        {
            var sharedStringItems = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ToList();

            if (!sharedStringItems.Any())
            {
                return val=>val;
            }

            return val => sharedStringItems[int.Parse(val)].InnerText;
        }
    }
}
