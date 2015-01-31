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
        private const int TimeOutAttemptsCount = 3;
        private const int MaxBatchSize = 1000;
        private const int MaxReadBatchSize = 5000;

        public GoogleDocGenerator(SpreadsheetsService service)
        {
            Service = service;
        }

        public async Task ExportToDocumentAsync<TModel>(string path, IReadOnlyList<ResGroupModel<TModel>> groups, IStatusProgress progress, CancellationToken cancellationToken) where TModel : IRowModel
        {
            progress.Report(StatusRes.GettingSpreadsheet);
            SpreadsheetFeed spreadSheets = await Service.GetFeedAsync(new SpreadsheetQuery(path), progress, cancellationToken);
            SpreadsheetEntry spreadsheet = (SpreadsheetEntry) spreadSheets.Entries.First();

            progress.Report(StatusRes.GettingWorksheets);
            WorksheetFeed worksheetsFeed = await Service.GetFeedAsync(spreadsheet.GetWorkSheetQuery(), progress, cancellationToken);

            int columsCount = groups.First().Tables.First().Rows.First().DataList.Count;
            List<KeyValuePair<string, WorksheetEntry>> newWorksheets = groups
                .Select(g => new KeyValuePair<string, WorksheetEntry>(g.GroupTitle, new WorksheetEntry((uint) (g.Tables.Sum(t => t.Rows.Count + 3) + /*Table name rows*/ g.Tables.Count - 1), (uint) columsCount, g.GroupTitle)))
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

                progress.Report(StatusRes.RenamingOldWorksheets);
                var progresses = progress.CreateParallelProgresses(worksheets2Rename.Count);
                //Renaming worksheets with matching names.
                await Task.WhenAll(worksheets2Rename.Zip(progresses, (ws, p) => Service.UpdateItemAsync(ws, p, cancellationToken)));
            }

            progress.Report(StatusRes.InsertingNewWorksheets);
            //Creating new worksheets.
            var insertingProgresses = progress.CreateParallelProgresses(newWorksheets.Count);
            var createdWorksheets = await Task.WhenAll(newWorksheets.Zip(insertingProgresses, (kvp, p) => Service.InsertItemAsync(spreadsheet.GetWorkSheetLink(), kvp.Value, p, cancellationToken)));
            newWorksheets = createdWorksheets.Select(ws => new KeyValuePair<string, WorksheetEntry>(ws.Title.Text, ws)).ToList();

            progress.Report(StatusRes.DeletingOldWorksheets);
            //Clearing of previous worksheets.
            var deletingProgresses = progress.CreateParallelProgresses(worksheetsFeed.Entries.Count);
            await Task.WhenAll(worksheetsFeed.Entries.Cast<WorksheetEntry>().Zip(deletingProgresses, (ws, p) => Service.DeleteItemAsync(ws, p, cancellationToken)).ToArray());

            progress.Report(StatusRes.PushingCells);
            var groupWorksheetsJoin = newWorksheets.Join(groups, ws => ws.Key, g => g.GroupTitle, (kvp, group) => new {Group = group, Worksheet = kvp.Value}).ToList();
            var groupProgresses = progress.CreateParallelProgresses(groupWorksheetsJoin.Select(j => (double) j.Worksheet.Cols*j.Worksheet.Rows).ToList());

            SemaphoreSlim semaphore = new SemaphoreSlim(MaxPushRequestsCount);
            await Task.WhenAll(groupWorksheetsJoin.Zip(groupProgresses, (j, p) => PushCellsAsync(j.Worksheet, j.Group, semaphore, p, cancellationToken)));
        }

        private async Task PushCellsAsync<TModel>(WorksheetEntry worksheet, ResGroupModel<TModel> group, SemaphoreSlim semaphore, IAggregateProgress progress, CancellationToken cancellationToken) 
            where TModel : IRowModel
        {
            var progresses = progress.CreateParallelProgresses(0.2, 0.8);
            var mappingProgress = progresses[0];
            var updatingProgress = progresses[1];



            var cellFeed = await Service.GetFeedAsync(worksheet.GetCellsQuery(), cancellationToken: cancellationToken);

            string cellsFeedUrl = cellFeed.Self;
            string batchUrl = cellFeed.Batch;

            var resCells = await CreateSheetCellInfos(group, cancellationToken);

            var cellGroups = Devide(resCells, MaxBatchSize);

            var mappingTasks = (await MapFeedBatchesAsync(cellGroups, cellsFeedUrl, batchUrl, mappingProgress, cancellationToken))
                .Select(async kvp =>
                {
                    var mappingFeed = kvp.Key;
                    var p = kvp.Value;

                    await semaphore.WaitAsync(cancellationToken);

                    try
                    {
                        return await Service.BatchFeedAsync(mappingFeed, TimeOutAttemptsCount, p, cancellationToken);
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
                    return await Service.BatchFeedAsync(updateFeed, TimeOutAttemptsCount, p, cancellationToken);
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

        #region Long running operations Tasks

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

        private Task<List<SheetCellInfo>> CreateSheetCellInfos<TModel>(ResGroupModel<TModel> group, CancellationToken cancellationToken)
            where TModel : IRowModel
        {
            return Task.Run(() =>
            {
                int rowIndex = 0;
                List<SheetCellInfo> resCells = new List<SheetCellInfo>();

                foreach (var table in group.Tables)
                {
                    resCells.Add(new SheetCellInfo(rowIndex, 0, table.TableTitle));
                    rowIndex += 2;

                    resCells.AddRange(table.Header.Columns.Select((header, colIndex) => new SheetCellInfo(rowIndex, colIndex, header.Title)));
                    rowIndex++;

                    foreach (var row in table.Rows)
                    {
                        resCells.AddRange(row.DataList.Select((dataCell, colIndex) => new SheetCellInfo(rowIndex, colIndex, dataCell.DataString)));

                        rowIndex++;
                    }

                    rowIndex++;
                    cancellationToken.ThrowIfCancellationRequested();
                }

                return resCells;
            }, cancellationToken);
        }

        private Task<ResGroupModel<TModel>> ProcessGroupAsync<TModel>(WorksheetEntry worksheet, IEnumerable<CellEntry> cells, CancellationToken cancellationToken) where TModel : IRowModel, new()
        {
            return Task.Run(() =>
            {
                List<string>[] rows = new List<string>[(int)worksheet.Rows];

                foreach (CellEntry cellEntry in cells)
                {
                    var cellsList = rows[cellEntry.Row - 1];
                    if (cellsList == null)
                    {
                        cellsList = new List<string>();
                        rows[cellEntry.Row - 1] = cellsList;
                    }

                    InsertElementAt(cellsList, (int) cellEntry.Column - 1, SheetCellInfo.GetValue(cellEntry));
                }

                var rowsEnumerator = rows.AsEnumerable().GetEnumerator();

                ResGroupModel<TModel> group = new ResGroupModel<TModel>
                {
                    GroupTitle = worksheet.Title.Text,
                };

                List<ResTableModel<TModel>> tables = new List<ResTableModel<TModel>>();

                int colsCount = (int) worksheet.Cols;

                List<string> row;
                while (rowsEnumerator.MoveNext())
                {
                    //Title row.
                    row = rowsEnumerator.Current;
                    if (row == null)
                    {
                        break;
                    }

                    var newTable = new ResTableModel<TModel> { TableTitle = row[0] };

                    //Space row.
                    rowsEnumerator.MoveNext();

                    //Header
                    rowsEnumerator.MoveNext();
                    row = rowsEnumerator.Current;

                    newTable.Header = new HeaderModel { Columns = row.Select(r => new ColumnModel { Title = r }).ToList() };

                    var rowsList = new List<RowModel<TModel>>();

                    while (rowsEnumerator.MoveNext())
                    {
                        row = rowsEnumerator.Current;
                        if (row == null)
                        {
                            break;
                        }

                        if (row.Count != colsCount)
                        {
                            //Adding empty cells.
                            for (int i = colsCount-1; i >= 0; i--)
                            {
                                var elem = row.ElementAtOrDefault(i);
                                if (elem == null)
                                {
                                    InsertElementAt(row, i, string.Empty);
                                    if (row.Count == colsCount)
                                    {
                                        break;
                                    }
                                }
                            }
                        }

                        rowsList.Add(new RowModel<TModel> {Model = new TModel {DataList = row.Select(c => new CellModel {Model = c}).ToList()}});
                    }

                    newTable.Rows = rowsList;
                    tables.Add(newTable);
                }

                group.Tables = tables;

                return group;
            }, cancellationToken);
        }

        #endregion

        public async Task<IReadOnlyList<ResGroupModel<TModel>>> ImportFromDocumentAsync<TModel>(string path, IStatusProgress progress, CancellationToken cancellationToken) where TModel : IRowModel, new()
        {
            progress.Report(StatusRes.GettingSpreadsheet);
            SpreadsheetFeed spreadSheets = await Service.GetFeedAsync(new SpreadsheetQuery(path), progress, cancellationToken);
            SpreadsheetEntry spreadsheet = (SpreadsheetEntry)spreadSheets.Entries.First();

            progress.Report(StatusRes.GettingWorksheets);
            WorksheetFeed worksheetsFeed = await Service.GetFeedAsync(spreadsheet.GetWorkSheetQuery(), progress, cancellationToken);

            List<WorksheetEntry> worksheets = worksheetsFeed.Entries.Cast<WorksheetEntry>().ToList();

            var progresses = progress.CreateParallelProgresses(worksheets.Select(ws => (double) ws.Cols*ws.Rows).ToList());

            progress.Report(StatusRes.GettingWorksheetsData);
            
            SemaphoreSlim semaphore = new SemaphoreSlim(MaxPushRequestsCount);
            var groupTasks = worksheets.Zip(progresses, (entry, p) => GetGroupAsync<TModel>(entry, semaphore, p, cancellationToken));

            return (await Task.WhenAll(groupTasks)).ToList();
        }

        private async Task<ResGroupModel<TModel>> GetGroupAsync<TModel>(WorksheetEntry worksheet, SemaphoreSlim semaphore, IAggregateProgress progress, CancellationToken cancellationToken) 
            where TModel : IRowModel, new()
        {
            uint rowsPerBatch = MaxReadBatchSize/worksheet.Cols;

            IEnumerable<uint> batchStartIndexes = Enumerable.Range(0, (int) (worksheet.Rows/rowsPerBatch + 1)).Select(batchIndex => (uint) (batchIndex*rowsPerBatch + 1));

            var cellsQueries = batchStartIndexes.Select(rowIndex =>
            {
                var cellsQuery = worksheet.GetCellsQuery();
                cellsQuery.MinimumRow = rowIndex;
                cellsQuery.MaximumRow = rowIndex + rowsPerBatch;
                if (cellsQuery.MaximumRow > worksheet.Rows)
                {
                    cellsQuery.MaximumRow = worksheet.Rows;
                }

                return cellsQuery;
            }).ToList();

            var progresses = progress.CreateParallelProgresses(cellsQueries.Select(cq => (double) cq.MaximumRow - cq.MinimumRow).ToList());

            var cellsTasks = cellsQueries.Zip(progresses, async (cellsQuery, p) =>
            {
                await semaphore.WaitAsync(cancellationToken);

                try
                {
                    var cellsFeed = await Service.GetFeedAsync(cellsQuery, p, cancellationToken);

                    return cellsFeed.Entries.Cast<CellEntry>();
                }
                finally
                {
                    semaphore.Release();
                }
            });


            var cells = (await Task.WhenAll(cellsTasks)).SelectMany(cs => cs).ToList();

            return await ProcessGroupAsync<TModel>(worksheet, cells, cancellationToken);
        }

        public async Task<IReadOnlyCollection<DocumentInfoModel>> GetDocuments()
        {
            var spreadSheets = await Service.GetFeedAsync(new SpreadsheetQuery());

            return spreadSheets.Entries.Cast<SpreadsheetEntry>()
                .Select(se => new DocumentInfoModel(se.SelfUri.Content, se.AlternateUri.Content, se.Title.Text))
                .ToList();
        }

        #region Extensions Methods

        private static List<List<TEntry>> Devide<TEntry>(IEnumerable<TEntry> enumerable, int range)
        {
            return enumerable.Select((e, index) => new KeyValuePair<int, TEntry>(index, e))
                .GroupBy(kvp => kvp.Key/range, kvp => kvp.Value)
                .Select(g => g.ToList()).ToList();
        }

        private static void InsertElementAt(List<string> list, int index, string element)
        {
            if (list.Count < index)
            {
                list.AddRange(Enumerable.Repeat(string.Empty, index - list.Count));
            }

            list.Insert(index, element);
        }

        #endregion
    }
}
