using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Common.Excel.Contracts;
using Common.Excel.GoogleSheets;
using Common.Excel.Models;
using Google.GData.Client;
using Google.GData.Spreadsheets;
using ResxPackage.Resources;

namespace Common.Excel.Implementation
{
    public class GoogleDocGenerator : IDocumentGenerator, IGoogleDocumentsService
    {
        public readonly SpreadsheetsService Service;
        private const int MaxPushRequestsCount = 5;
        private const int MaxBatchSize = 1000;

        public GoogleDocGenerator(SpreadsheetsService service)
        {
            Service = service;
        }

        public async Task ExportToDocumentAsync<TModel>(string path, IReadOnlyList<ResGroupModel<TModel>> groups, IStatusProgress progress, CancellationToken cancellationToken) where TModel : IRowModel
        {
            progress.Report(StatusRes.GettingSpreadsheet, 0);
            SpreadsheetFeed spreadSheets = await Service.GetFeedAsync(new SpreadsheetQuery(path), progress, cancellationToken);
            SpreadsheetEntry spreadsheet = (SpreadsheetEntry) spreadSheets.Entries.First();

            progress.Report(StatusRes.GettingWorksheets, 0);
            WorksheetFeed worksheetsFeed = await Service.GetFeedAsync(spreadsheet.GetWorkSheetQuery(), progress, cancellationToken);

            int columsCount = groups.First().Tables.First().Rows.First().DataList.Count;
            List<KeyValuePair<string, WorksheetEntry>> newWorksheets = groups
                .Select(g => new KeyValuePair<string, WorksheetEntry>(g.GroupTitle, new WorksheetEntry((uint) (g.Tables.Sum(t => t.Rows.Count + 2) + g.Tables.Count - 1), (uint) columsCount, g.GroupTitle)))
                .ToList();

            List<WorksheetEntry> worksheets2Rename = worksheetsFeed.Entries.Cast<WorksheetEntry>()
                .Where(ws => newWorksheets.Any(nws => String.Equals(nws.Key, ws.Title.Text, StringComparison.OrdinalIgnoreCase)))
                .ToList();
            if (worksheets2Rename.Count != 0)
            {
                foreach (var worksheetEntry in worksheets2Rename)
                {
                    worksheetEntry.Title.Text = Guid.NewGuid().ToString("N");
                    worksheetEntry.BatchData = new GDataBatchEntryData(GDataBatchOperationType.update);
                }

                progress.Report(StatusRes.RenamingOldWorksheets, 0);
                var progresses = progress.CreateParallelProgresses(worksheets2Rename.Count);
                //Renaming worksheets with matching names.
                await Task.WhenAll(worksheets2Rename.Zip(progresses, (ws, p) => Service.UpdateItemAsync(ws, p, cancellationToken)));
            }

            progress.Report(StatusRes.InsertingNewWorksheets, 0);
            //Creating new worksheets.
            var insertingProgresses = progress.CreateParallelProgresses(newWorksheets.Count);
            var createdWorksheets = await Task.WhenAll(newWorksheets.Zip(insertingProgresses, (kvp, p) => Service.InsertItemAsync(spreadsheet.GetWorkSheetLink(), kvp.Value, p, cancellationToken)));
            newWorksheets = createdWorksheets.Select(ws => new KeyValuePair<string, WorksheetEntry>(ws.Title.Text, ws)).ToList();

            progress.Report(StatusRes.DeletingOldWorksheets, 0);
            //Clearing of previous worksheets.
            var deletingProgresses = progress.CreateParallelProgresses(worksheetsFeed.Entries.Count);
            await Task.WhenAll(worksheetsFeed.Entries.Cast<WorksheetEntry>().Zip(deletingProgresses, (ws, p) => Service.DeleteItemAsync(ws, p, cancellationToken)).ToArray());

            progress.Report(StatusRes.PushingCells, 0);
            var groupWorksheetsJoin = newWorksheets.Join(groups, ws => ws.Key, g => g.GroupTitle, (kvp, group) => new {Group = group, Worksheet = kvp.Value}).ToList();
            var groupProgresses = progress.CreateParallelProgresses(groupWorksheetsJoin.Select(j => (double) j.Worksheet.Cols*j.Worksheet.Rows).ToList());

            SemaphoreSlim semaphore = new SemaphoreSlim(MaxPushRequestsCount);
            await Task.WhenAll(groupWorksheetsJoin.Zip(groupProgresses, (j, p) => PushCellsAsync(j.Worksheet, j.Group, semaphore, p, cancellationToken)));
        }

        private async Task PushCellsAsync<TModel>(WorksheetEntry worksheet, ResGroupModel<TModel> group, SemaphoreSlim semaphore, IAggregateProgress progress, CancellationToken cancellationToken) 
            where TModel : IRowModel
        {
            int rowIndex = 0;

            var progresses = progress.CreateParallelProgresses(0.3, 0.7);
            var mappingProgress = progresses[0];
            var updatingProgress = progresses[1];

            List<SheetCellInfo> resCells = new List<SheetCellInfo>();

            var cellFeed = await Service.GetFeedAsync(worksheet.GetCellsQuery(), cancellationToken: cancellationToken);

            foreach (var table in group.Tables)
            {
                resCells.Add(new SheetCellInfo(rowIndex, 0, table.TableTitle));
                rowIndex += 2;

                foreach (var row in table.Rows)
                {
                    resCells.AddRange(row.DataList.Select((dataCell, colIndex) => new SheetCellInfo(rowIndex, colIndex, dataCell.DataString)));

                    rowIndex++;
                }
            }

            string cellsFeedUrl = cellFeed.Self;
            string batchUrl = cellFeed.Batch;

            var cellGroups = Devide(resCells, MaxBatchSize);

            var mappingTasks = (await MapFeedBatchesAsync(cellGroups, cellsFeedUrl, batchUrl, mappingProgress, cancellationToken))
                .Select(async kvp =>
                {
                    var mappingFeed = kvp.Key;
                    var p = kvp.Value;

                    await semaphore.WaitAsync(cancellationToken);

                    try
                    {
                        return await Service.BatchFeedAsync(mappingFeed, p, cancellationToken);
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                });

            //Mapping cells.
            List<CellFeed> mappingFeeds = new List<CellFeed>(cellGroups.Count);
            foreach (var mappingTask in mappingTasks)
            {
                mappingFeeds.Add(await mappingTask);
            }

            //Updating cells.
            var updatingTasks = (await UpdateFeedBatchesAsync(cellGroups, mappingFeeds, cellsFeedUrl, batchUrl, updatingProgress, cancellationToken))
            .Select(async kvp =>
            {
                var updateFeed = kvp.Key;
                var p = kvp.Value;

                await semaphore.WaitAsync(cancellationToken);

                try
                {
                    return await Service.BatchFeedAsync(updateFeed, p, cancellationToken);
                }
                finally
                {
                    semaphore.Release();
                }
            });

            foreach (var updatingTask in updatingTasks)
            {
                await updatingTask;
            }
        }

        private Task<IReadOnlyCollection<KeyValuePair<CellFeed, IAggregateProgress>>> MapFeedBatchesAsync(List<List<SheetCellInfo>> cellGroups, string cellsFeedUrl, string batchUrl, IAggregateProgress progress, CancellationToken cancellationToken)
        {
            return Task.Run<IReadOnlyCollection<KeyValuePair<CellFeed, IAggregateProgress>>>(() =>
            {
                var progresses = progress.CreateParallelProgresses(cellGroups.Select(g => (double) g.Count).ToList());

                var batches = cellGroups.Zip(progresses, (cells, p) =>
                {
                    var mappingFeed = new CellFeed(new Uri(cellsFeedUrl), Service) {Batch = batchUrl};

                    foreach (var cellEntry in cells)
                    {
                        mappingFeed.Entries.Add(cellEntry.BatchCreateEntry(cellsFeedUrl));
                    }

                    cancellationToken.ThrowIfCancellationRequested();

                    return new KeyValuePair<CellFeed, IAggregateProgress>(mappingFeed, p);
                });

                return batches.ToList();
            }, cancellationToken);
        }

        private Task<IReadOnlyCollection<KeyValuePair<CellFeed, IAggregateProgress>>> UpdateFeedBatchesAsync(
            List<List<SheetCellInfo>> cellGroups, 
            IEnumerable<CellFeed> mappingFeeds, 
            string cellsFeedUrl, 
            string batchUrl, 
            IAggregateProgress progress, 
            CancellationToken cancellationToken)
        {
            return Task.Run<IReadOnlyCollection<KeyValuePair<CellFeed, IAggregateProgress>>>(
                () =>
                {
                    var cellsJoin = mappingFeeds.Zip(cellGroups, (mappingFeed, cellsGroup) =>
                        from feedCell in mappingFeed.Entries.Cast<CellEntry>()
                        from cell in cellsGroup
                        where feedCell.Row == cell.Row && feedCell.Column == cell.Column
                        select new {Cell = cell, FeedCell = feedCell});
                    var updatingProgresses = progress.CreateParallelProgresses(cellGroups.Select(g => (double) g.Count).ToList());

                    var batches = cellsJoin.Zip(updatingProgresses, (cells, p) =>
                    {
                        var updateFeed = new CellFeed(new Uri(cellsFeedUrl), Service) {Batch = batchUrl};

                        foreach (var pair in cells)
                        {
                            updateFeed.Entries.Add(pair.Cell.BatchUpdateEntry(pair.FeedCell));
                        }

                        cancellationToken.ThrowIfCancellationRequested();

                        return new KeyValuePair<CellFeed, IAggregateProgress>(updateFeed, p);
                    });

                    return batches.ToList();
                }, cancellationToken);
        }

        public Task<IReadOnlyList<ResGroupModel<TModel>>> ImportFromExcelAsync<TModel>(string path, IStatusProgress progress, CancellationToken cancellationToken) where TModel : IRowModel, new()
        {
            throw new NotImplementedException();
        }

        public async Task<IReadOnlyCollection<DocumentInfoModel>> GetDocuments()
        {
            var spreadSheets = await Service.GetFeedAsync(new SpreadsheetQuery());

            return spreadSheets.Entries.Cast<SpreadsheetEntry>()
                .Select(se => new DocumentInfoModel(se.SelfUri.Content, se.AlternateUri.Content, se.Title.Text))
                .ToList();
        }

        private static List<List<TEntry>> Devide<TEntry>(IEnumerable<TEntry> enumerable, int range)
        {
            return enumerable.Select((e, index) => new KeyValuePair<int, TEntry>(index, e))
                .GroupBy(kvp => kvp.Key/range, kvp => kvp.Value)
                .Select(g => g.ToList()).ToList();
        }
    }
}
