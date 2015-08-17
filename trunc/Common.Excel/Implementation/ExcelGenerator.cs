using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Common.Excel.Contracts;
using Common.Excel.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using ResxPackage.Resources;

namespace Common.Excel.Implementation
{
    public class ExcelGenerator : IDocumentGenerator
    {
        public Task ExportToDocumentAsync<TModel>(string path, IReadOnlyList<ResGroupModel<TModel>> groups, IStatusProgress progress, CancellationToken cancellationToken) where TModel : IRowModel
        {
            return Task.Run(()=>ExportToDocument(path, groups, progress, cancellationToken), cancellationToken);
        }

        public Task<IReadOnlyList<ResGroupModel<TModel>>> ImportFromDocumentAsync<TModel>(string path, IStatusProgress progress, CancellationToken cancellationToken) where TModel : IRowModel, new()
        {
            return Task.Run(()=>ImportFromDocument<TModel>(path), cancellationToken);
        }

        private void ExportToDocument<TModel>(string path, IReadOnlyList<ResGroupModel<TModel>> groups, IStatusProgress progress, CancellationToken cancellationToken) where TModel : IRowModel
        {
            if (groups.Count == 0)
            {
                throw new ArgumentException(ErrorsRes.TherIsNotFilesToExport, "groups");
            }

            progress.Report(StatusRes.ExportingToExcel);

            using (ExcelPackage package = new ExcelPackage())
            {
                /* Create the worksheet. */
                var workbook = package.Workbook;

                double totalRows = groups.Sum(g => g.Tables.Sum(t => t.Rows.Count));
                int rowIndexReport = 1;

                foreach (var group in groups)
                {
                    var worksheet = workbook.Worksheets.Add(group.GroupTitle);
                    ExcelStyle defaultStyle = worksheet.Cells[1, 1].Style;

                    int rowIndex = 1;

                    int columnsCount = groups.SelectMany(g => g.Tables.Select(t => t.Header.Columns.Count)).Max();

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

                        worksheet.Column(i + 1).Width = (int)GetDefaultFontWidth(defaultStyle, longestString);
                    }

                    //Setting Columns
                    foreach (var resTableModel in group.Tables)
                    {
                        var headerCell = worksheet.Cells[rowIndex++, 1];

                        headerCell.Style.Font.Bold = true;
                        headerCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        headerCell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                        headerCell.Value = resTableModel.TableTitle;

                        rowIndex++;

                        for (int columnIndex = 0; columnIndex < resTableModel.Header.Columns.Count; columnIndex++)
                        {
                            var languageHeader = worksheet.Cells[rowIndex, columnIndex + 1];

                            languageHeader.Style.Font.Bold = true;
                            languageHeader.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            languageHeader.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                            languageHeader.Value = resTableModel.Header.Columns[columnIndex].Title;
                        }

                        rowIndex++;

                        foreach (var tableRow in resTableModel.Rows)
                        {
                            for (int columnIndex = 0; columnIndex < tableRow.DataList.Count; columnIndex++)
                            {
                                var cell = tableRow.DataList[columnIndex];

                                var valueCell = worksheet.Cells[rowIndex, columnIndex + 1];
                                valueCell.Value = cell.DataString;

                                if (cell.Hilight)
                                {
                                    valueCell.Style.Font.Color.SetColor(Color.Red);
                                }
                            }

                            rowIndex++;

                            rowIndexReport++;
                        }

                        progress.Report(100 * rowIndexReport / totalRows);
                        cancellationToken.ThrowIfCancellationRequested();

                        rowIndex++;
                    }
                }

                package.SaveAs(new FileInfo(path));
            }
        }

        private IReadOnlyList<ResGroupModel<TModel>> ImportFromDocument<TModel>(string path) where TModel : IRowModel, new()
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(path)))
            {
                var workbook = package.Workbook;

                var dataList = new List<ResGroupModel<TModel>>();

                foreach (ExcelWorksheet worksheetPart in workbook.Worksheets)
                {
                    var tables = new List<ResTableModel<TModel>>();
                    var newGroup = new ResGroupModel<TModel> { GroupTitle = worksheetPart.Name, Tables = tables };

                    var cells = worksheetPart.Cells.Select(cell => new {cell.Start.Row, cell.Start.Column, Value = (string) cell.Value}).ToList();
                    int maxRow = cells.Max(c => c.Row);

                    string[][] cellsArray = new string[maxRow][];

                    foreach (var row in cells.GroupBy(c => c.Row))
                    {
                        cellsArray[row.Key - 1] = row.OrderBy(c => c.Column).Select(c => c.Value).ToArray();
                    }

                    for (int i = 0; i < cellsArray.Length; i++)
                    {
                        var table = new ResTableModel<TModel>();

                        table.TableTitle = cellsArray[i][0];

                        i += 2;
                        table.Header = new HeaderModel {Columns = cellsArray[i].Select(c => new ColumnModel {Title = c}).ToList()};

                        var rows = new List<RowModel<TModel>>();
                        string[] rowsArray;
                        while (++i < cellsArray.Length && (rowsArray = cellsArray[i]) != null)
                        {
                            rows.Add(new RowModel<TModel>
                            {
                                Model = new TModel
                                {
                                    DataList = rowsArray.Select(a => new CellModel {DataString = a}).ToList()
                                }
                            });
                        }

                        table.Rows = rows;
                    }

                    dataList.Add(newGroup);
                }

                return dataList;
            }
        }

        public static double GetDefaultFontWidth(ExcelStyle style, string text)
        {
            string font = style.Font.Name;
            int fontSize = (int)style.Font.Size;
            Font stringFont = new Font(font, fontSize);
            return GetWidth(stringFont, text) + 2.0;
        }

        private static double GetWidth(Font stringFont, string text)
        {
            // This formula is based on this article plus a nudge ( + 0.2M )
            // http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.column.width.aspx
            // Truncate(((256 * Solve_For_This + Truncate(128 / 7)) / 256) * 7) = DeterminePixelsOfString

            Size textSize = TextRenderer.MeasureText(text, stringFont, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.SingleLine | TextFormatFlags.LeftAndRightPadding);
            double width = (((textSize.Width / (double)7) * 256) - ((double)128 / 7)) / 256;
            width = (double)decimal.Round((decimal)width + 0.2M, 2);

            return width;
        }
    }
}
