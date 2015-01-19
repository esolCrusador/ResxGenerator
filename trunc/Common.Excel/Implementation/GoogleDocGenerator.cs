using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Threading.Tasks;
using Common.Excel.Contracts;
using Common.Excel.GoogleSheets;
using Common.Excel.Models;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office2010.Word;
using Google.GData.Client;
using Google.GData.Spreadsheets;

namespace Common.Excel.Implementation
{
    public class GoogleDocGenerator : IDocumentGenerator, IGoogleDocumentsService
    {
        public readonly SpreadsheetsService Service;

        public GoogleDocGenerator(SpreadsheetsService service)
        {
            Service = service;
        }

        public async Task ExportToDocumentAsync<TModel>(string path, IReadOnlyList<ResGroupModel<TModel>> groups) where TModel : IRowModel
        {
            SpreadsheetFeed spreadSheets = await Service.GetFeedAsync(new SpreadsheetQuery(path), new object());
            SpreadsheetEntry spreadsheet = (SpreadsheetEntry) spreadSheets.Entries.First();

            WorksheetFeed worksheetsFeed = await Service.GetFeedAsync(spreadsheet.GetWorkSheetQuery(), new object());

            int columsCount = groups.First().Tables.First().Rows.First().DataList.Count;
            IEnumerable<KeyValuePair<string, WorksheetEntry>> worksheets = groups
                .Select(g => new KeyValuePair<string, WorksheetEntry>(g.GroupTitle, new WorksheetEntry((uint)(g.Tables.Sum(t => t.Rows.Count + 2) + g.Tables.Count-1), (uint)columsCount, g.GroupTitle)))
                .ToList();

            var worksheets2Rename = worksheetsFeed.Entries.Cast<WorksheetEntry>()
                .Where(ws => worksheets.Any(nws => String.Equals(nws.Key, ws.Title.Text, StringComparison.OrdinalIgnoreCase)))
                .ToList();
            if (worksheets2Rename.Count != 0)
            {
                foreach (var worksheetEntry in worksheets2Rename)
                {
                    worksheetEntry.Title.Text = Guid.NewGuid().ToString("N");
                    worksheetEntry.BatchData = new GDataBatchEntryData(GDataBatchOperationType.update);
                }

                await Task.WhenAll(worksheets2Rename.Select(ws => Service.UpdateItemAsync(ws, new object())));
            }

            //Creating new worksheets.
            var createdWorksheets = await Task.WhenAll(worksheets.Select(kvp => Service.InsertItemAsync(spreadsheet.GetWorkSheetLink(), kvp.Value, new object())));
            worksheets = createdWorksheets.Select(ws => new KeyValuePair<string, WorksheetEntry>(ws.Title.Text, ws)).ToList();

            //Clearing of previous worksheets.
            await Task.WhenAll(worksheetsFeed.Entries.Cast<WorksheetEntry>().Select(ws => Service.DeleteItemAsync(ws, new object())).ToArray());

            var groupWorksheetsJoin = worksheets.Join(groups, ws => ws.Key, g => g.GroupTitle, (kvp, group) => new {Group = group, Worksheet = kvp.Value}).ToList();

            await Task.WhenAll(groupWorksheetsJoin.Select(j => PushCellsAsync(j.Worksheet, j.Group)));
        }

        private async Task PushCellsAsync<TModel>(WorksheetEntry worksheet, ResGroupModel<TModel> group) 
            where TModel : IRowModel
        {
            int rowIndex = 0;
            CellFeed cellFeed;
            CellFeed pushFeed;

            List<SheetCellInfo> cells = new List<SheetCellInfo>();

            cellFeed = await Service.GetFeedAsync(worksheet.GetCellsQuery(), new object());

            foreach (var table in group.Tables)
            {
                cells.Add(new SheetCellInfo(rowIndex, 0, table.TableTitle));
                rowIndex += 2;

                foreach (var row in table.Rows)
                {
                    cells.AddRange(row.DataList.Select((dataCell, colIndex) => new SheetCellInfo(rowIndex, colIndex, dataCell.DataString)));

                    rowIndex++;
                }
            }

            string cellsFeedUrl = cellFeed.Self;
            string batchUrl = cellFeed.Batch;
            pushFeed = new CellFeed(new Uri(cellsFeedUrl), Service) { Batch = batchUrl };
            foreach (var cellEntry in cells)
            {
                pushFeed.Entries.Add(cellEntry.BatchCreateEntry(cellsFeedUrl));
            }

            cellFeed = await Service.BatchFeedAsync(pushFeed, new object());

            pushFeed = new CellFeed(new Uri(cellsFeedUrl), Service) { Batch = batchUrl };
            var cellsJoin = from feedCell in cellFeed.Entries.Cast<CellEntry>()
                            from cell in cells
                            where feedCell.Row == cell.Row && feedCell.Column == cell.Column
                            select new { Cell = cell, FeedCell = feedCell };
            foreach (var pair in cellsJoin)
            {
                pushFeed.Entries.Add(pair.Cell.BatchUpdateEntry(pair.FeedCell));
            }

            cellFeed = await Service.BatchFeedAsync(pushFeed, new object());
        }

        public Task<IReadOnlyList<ResGroupModel<TModel>>> ImportFromExcelAsync<TModel>(string path) where TModel : IRowModel, new()
        {
            throw new NotImplementedException();
        }

        public async Task<IReadOnlyCollection<DocumentInfoModel>> GetDocuments()
        {
            var spreadSheets = await Service.GetFeedAsync(new SpreadsheetQuery(), new object());

            return spreadSheets.Entries.Cast<SpreadsheetEntry>()
                .Select(se => new DocumentInfoModel(se.SelfUri.Content, se.AlternateUri.Content, se.Title.Text))
                .ToList();
        }
    }
}
