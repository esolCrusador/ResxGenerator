using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Common.Excel.Contracts;
using Common.Excel.Models;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MoreLinq;

namespace Common.Excel.Implementation
{
    public class ExcelGenerator : IExcelGenerator
    {
        //For Excel2007 and above .xlsx files
        const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        string GetFileName(string title)
        {
            return string.Format("{0}.xlsx", title);
        }

        public FileInfoContainer ExportToExcel(ExcelExportModel mdl)
        {
            using (var stream = new MemoryStream())
            {
                /* Create the worksheet. */
                SpreadsheetDocument spreadsheet = Excel.CreateWorkbook(stream);
                Excel.AddBasicStyles(spreadsheet);
                Excel.AddAdditionalStyles(spreadsheet);
                Excel.AddWorksheet(spreadsheet, mdl.Title);
                Worksheet worksheet = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet;

                /* Add the column titles to the worksheet. */
                for (int colId = 0; colId < mdl.ColumnHeaders.Count; colId++)
                {
                    // If the column has a title, use it.  Otherwise, use the field name.
                    Excel.SetColumnHeadingValue(spreadsheet, worksheet, Convert.ToUInt32(colId + 1), 1,
                        mdl.ColumnHeaders[colId],
                        false, false);

                    // Is there are column width defined?
                    var allRows = mdl.Rows.Select(r => r[colId]).Concat(new[] {mdl.ColumnHeaders[colId]});
                    var maxString = allRows.MaxBy(c=>c.Length);
                    var width = Excel.GetDefaultFontWidth(maxString);
                    if (width>0)
                    {
                        Excel.SetColumnWidth(worksheet, colId + 1, (int)width);
                    }
                }
                
                // For each row of data...
                for (int rowId = 0; rowId < mdl.Rows.Count; rowId++)
                {
                    for (int colId = 0; colId < mdl.ColumnHeaders.Count; colId++)
                    {
                        // Set the field value in the spreadsheet for the current row and column.
                        Excel.SetCellValue(spreadsheet, worksheet, Convert.ToUInt32(colId + 1), Convert.ToUInt32(rowId + 2),
                            mdl.Rows[rowId][colId],
                            false, false);
                    }
                }
                
                //Save the worksheet
                worksheet.Save();
                spreadsheet.Close();
                return new FileInfoContainer(stream.ToArray(), GetFileName(mdl.Title));
            }
        }

        public FileInfoContainer ExportToExcel<TModel>(IReadOnlyList<ResGroupModel<TModel>> groups, string title) where TModel : IRowModel
        {
            if (groups.Count == 0)
            {
                throw new ArgumentException("There is not resource files to export", "groups");
            }

            using (var stream = new MemoryStream())
            {
                /* Create the worksheet. */
                SpreadsheetDocument spreadsheet = Excel.CreateWorkbook(stream);
                Excel.AddBasicStyles(spreadsheet);
                Excel.AddAdditionalStyles(spreadsheet);
                Excel.AddWorksheet(spreadsheet, title);
                Worksheet worksheet = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet;

                
                uint rowIndex = 1;

                int columnsCount = groups.SelectMany(g => g.Tables.Select(t => t.Header.Columns.Count)).Max();
                List<int> columnWidthes = new List<int>(columnsCount);

                for (int i = 0; i < columnsCount; i++)
                {
                    string longestString = groups.SelectMany(g =>
                        g.Tables.Select(t =>
                        {
                            string colTitle = t.Header.Columns[i].Title;
                            string longestRow = t.Rows.Select(r => r.DataList[i].DataString).MaxBy(str => str.Length);

                            return colTitle.Length > longestRow.Length ? colTitle : longestRow;
                        })
                        ).MaxBy(str => str.Length);

                    columnWidthes.Add((int)Excel.GetDefaultFontWidth(longestString));
                }

                foreach (var @group in groups)
                {
                    Excel.SetColumnHeadingValue(spreadsheet, worksheet, 1, rowIndex++, @group.GroupTitle, false, false);
                    Excel.SetCellValue(spreadsheet, worksheet, 1, rowIndex++, " ", false, false);

                    //Setting Columns
                    foreach (var resTableModel in @group.Tables)
                    {
                        Excel.SetColumnHeadingValue(spreadsheet, worksheet, 1, rowIndex++, resTableModel.TableTitle, false, false);
                        Excel.SetCellValue(spreadsheet, worksheet, 1, rowIndex++, " ", false, false);

                        for (int columnIndex = 0; columnIndex < resTableModel.Header.Columns.Count; columnIndex++)
                        {
                            Excel.SetColumnHeadingValue(spreadsheet, worksheet, (uint)columnIndex + 1, rowIndex, resTableModel.Header.Columns[columnIndex].Title, false, false);

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

                                Excel.SetCellValue(spreadsheet, worksheet, (uint)columnIndex + 1, rowIndex, cell.DataString, false, false, highlight: cell.Hilight);
                            }

                            rowIndex++;
                        }

                        Excel.SetCellValue(spreadsheet, worksheet, 1, rowIndex++, " ", false, false);
                    }
                }

                worksheet.Save();
                spreadsheet.Close();
                return new FileInfoContainer(stream.ToArray(), GetFileName(title));
            }
        }

        public IReadOnlyList<ResGroupModel<TModel>> ImportFromExcel<TModel>(FileInfoContainer file) where TModel : IRowModel, new()
        {
            using (MemoryStream stream = new MemoryStream(file.Bytes))
            {
                SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, false);
                WorkbookPart workbookPart = spreadsheet.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                var itemGetter = GetSharedStringItemById(workbookPart);

                List<List<string>> data = sheetData.Cast<Row>()
                    .Select(row => row.ChildElements.Cast<Cell>()
                        .Select(cell => cell.CellValue != null ? itemGetter(cell.CellValue.Text) : null)
                        .ToList()
                    )
                    .ToList();

                var dataList = new List<ResGroupModel<TModel>>();

                var dataEnumerator  = data.GetEnumerator();
                List<string> rowCells = null;

                Func<bool> moveNextFunc = () =>
                    dataEnumerator.MoveNext() && (rowCells = dataEnumerator.Current) != null && rowCells.Count != 0 && !String.IsNullOrWhiteSpace(rowCells[0]);
                
                while (moveNextFunc())
                {
                    var newGroup = new ResGroupModel<TModel> {GroupTitle = rowCells[0]};

                    var tables = new List<ResTableModel<TModel>>();

                    dataEnumerator.MoveNext();

                    while (moveNextFunc())
                    {
                        var newTable = new ResTableModel<TModel>
                        {
                            TableTitle = rowCells[0]
                        };

                        dataEnumerator.MoveNext();
                        moveNextFunc();

                        newTable.Header = new HeaderModel<TModel>
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

                    //Space between projects
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

        #region just other possible variations
        //public static Stream ExportToExcel<T>(List<T> data, List<ColumnModel> columnsMetadatas, string title)
        //{
        //    var stream = new MemoryStream();
        //    // Create the worksheet.
        //    SpreadsheetDocument spreadsheet = Excel.CreateWorkbook(stream);
        //    Excel.AddBasicStyles(spreadsheet);
        //    Excel.AddAdditionalStyles(spreadsheet);
        //    Excel.AddWorksheet(spreadsheet, title);
        //    Worksheet worksheet = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet;

        //    //Add the column titles to the worksheet.
        //    for (var i = 0; i < columnsMetadatas.Count; i++)
        //    {
        //        // If the column has a title, use it.  Otherwise, use the field name.
        //        Excel.SetColumnHeadingValue(spreadsheet, worksheet, Convert.ToUInt32(i + 1),
        //            (string.IsNullOrWhiteSpace(columnsMetadatas[i].Title))
        //                ? columnsMetadatas[i].Field
        //                : columnsMetadatas[i].Title,
        //            false, false);

        //        // Is there are column width defined?
        //        //Excel.SetColumnWidth(worksheet, i + 1, columnsMetadatas[i].width != null
        //        //    ? int.Parse(LeadingInteger.Match(columnsMetadatas[i].width.ToString()).Value) / 4
        //        //    : 25);
        //    }

        //    var vp = new DataValueProvider(typeof(T));

        //    //Add the data to the worksheet.
        //    for (int rowId = 0; rowId < data.Count; rowId++)
        //    {
        //        //for each column...
        //        for (var columnId = 0; columnId < columnsMetadatas.Count; columnId++)
        //        {
        //            var fieldName = columnsMetadatas[columnId].Field;
        //            // Set the field value in the spreadsheet for the current row and column.
        //            Excel.SetCellValue(spreadsheet, worksheet, Convert.ToUInt32(columnId + 1),
        //                Convert.ToUInt32(rowId + 2),
        //                vp.GetValue(data[rowId], fieldName),
        //                false, false);
        //        }
        //    }

        //    worksheet.Save();
        //    spreadsheet.Close();
        //    return stream;
        //}

        //// http://stackoverflow.com/questions/975455/is-there-an-equivalent-to-javascript-parseint-in-c
        //private static readonly Regex LeadingInteger = new Regex(@"^(-?\d+)");

        //public static FileInfoContainer ExportToExcel(dynamic data, dynamic metadata, string title)
        //{
        //    using (var stream = new MemoryStream())
        //    {
        //        /* Create the worksheet. */

        //        SpreadsheetDocument spreadsheet = Excel.CreateWorkbook(stream);
        //        Excel.AddBasicStyles(spreadsheet);
        //        Excel.AddAdditionalStyles(spreadsheet);
        //        Excel.AddWorksheet(spreadsheet, title);
        //        Worksheet worksheet = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet;

        //        /* Add the column titles to the worksheet. */
        //        for (int mdx = 0; mdx < metadata.Count; mdx++)
        //        {
        //            // If the column has a title, use it.  Otherwise, use the field name.
        //            Excel.SetColumnHeadingValue(spreadsheet, worksheet, Convert.ToUInt32(mdx + 1),
        //                (metadata[mdx].title == null || metadata[mdx].title == "&nbsp;")
        //                    ? metadata[mdx].field.ToString()
        //                    : metadata[mdx].title.ToString(),
        //                false, false);

        //            // Is there are column width defined?
        //            Excel.SetColumnWidth(worksheet, mdx + 1, metadata[mdx].width != null
        //                ? int.Parse(LeadingInteger.Match(metadata[mdx].width.ToString()).Value) / 4
        //                : 25);
        //        }


        //        // For each row of data...
        //        for (int idx = 0; idx < data.Count; idx++)
        //        {
        //            // For each column...
        //            for (int mdx = 0; mdx < metadata.Count; mdx++)
        //            {
        //                // Set the field value in the spreadsheet for the current row and column.
        //                Excel.SetCellValue(spreadsheet, worksheet, Convert.ToUInt32(mdx + 1), Convert.ToUInt32(idx + 2),
        //                    data[idx][metadata[mdx].field.ToString()].ToString(),
        //                    false, false);
        //            }
        //        }


        //        /* Save the worksheet and store it in Session using the spreadsheet title. */

        //        worksheet.Save();
        //        spreadsheet.Close();
        //        return new FileInfoContainer(stream.ToArray(), ExcelGenerator.GetFileName(title), ContentType);
        //    }
        //}

        #endregion
    }
}
