using System;
using System.Threading;
using System.Threading.Tasks;
using Google.GData.Client;
using Google.GData.Spreadsheets;

namespace Common.Excel.GoogleSheets
{
    public static class SpreadSheetServiceExtensions
    {
        #region Async Feed Overloads

        public static Task<SpreadsheetFeed> GetFeedAsync(this SpreadsheetsService service, SpreadsheetQuery query, IAggregateProgress progress = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            return GetFeedAsync<SpreadsheetFeed, SpreadsheetQuery>(service, query, progress, cancellationToken);
        }

        public static Task<WorksheetFeed> GetFeedAsync(this SpreadsheetsService service, WorksheetQuery query, IAggregateProgress progress = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            return GetFeedAsync<WorksheetFeed, WorksheetQuery>(service, query, progress, cancellationToken);
        }

        public static Task<CellFeed> GetFeedAsync(this SpreadsheetsService service, CellQuery query, IAggregateProgress progress = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            return GetFeedAsync<CellFeed, CellQuery>(service, query, progress, cancellationToken);
        }

        public static Task<ListFeed> GetFeedAsync(this SpreadsheetsService service, ListQuery query, IAggregateProgress progress = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            return GetFeedAsync<ListFeed, ListQuery>(service, query, progress, cancellationToken);
        }

        private static Task<TResult> GetFeedAsync<TResult, TQuery>(SpreadsheetsService service, TQuery query, IAggregateProgress progress, CancellationToken cancellationToken)
            where TQuery : FeedQuery
            where TResult : AtomFeed
        {
            var userData = new object();

            service.QueryFeedAync(query.Uri, query.ModifiedSince, userData);

            return QueryFeedAsync<TResult>(service, userData, progress, cancellationToken);
        }

        #endregion 

        #region Async Batch

        public static Task<TFeed> BatchFeedAsync<TFeed>(this SpreadsheetsService service, TFeed feed, IAggregateProgress progress = null, CancellationToken cancellationToken = default (CancellationToken))
            where TFeed: AtomFeed
        {
            var userData = new object();

           
            service.BatchAsync(feed, new Uri(feed.Batch), userData);

            return QueryFeedAsync<TFeed>(service, userData, progress, cancellationToken);
        }

        #endregion

        #region Async Entry Operations

        public static Task DeleteItemAsync<TEntry>(this SpreadsheetsService service, TEntry item, IAggregateProgress progress = null, CancellationToken cancellationToken = default (CancellationToken))
            where  TEntry: AtomEntry
        {
            var userData = new object();

            service.DeleteAsync(item, true, userData);

            return QueryEntryAsync<TEntry>(service, userData, progress, cancellationToken);
        }

        public static Task<TEntry> InsertItemAsync<TEntry>(this SpreadsheetsService service, string feedUri, TEntry item, IAggregateProgress progress = null, CancellationToken cancellationToken = default (CancellationToken))
            where TEntry: AtomEntry
        {
            var userData = new object();

            service.InsertAsync(new Uri(feedUri), item, userData);

            return QueryEntryAsync<TEntry>(service, userData, progress, cancellationToken);
        }

        public static Task<TEntry> UpdateItemAsync<TEntry>(this SpreadsheetsService service, TEntry item, IAggregateProgress progress=null, CancellationToken cancellationToken = default (CancellationToken))
            where TEntry : AtomEntry
        {
            var userData = new object();

            service.UpdateAsync(item, userData);

            return QueryEntryAsync<TEntry>(service, userData, progress, cancellationToken);
        }

        #endregion

        #region Event To Task conversion

        private static Task<TResult> QueryFeedAsync<TResult>(SpreadsheetsService service, object userData, IAggregateProgress progress, CancellationToken cancellationToken)
            where TResult: AtomFeed
        {
            return ExecuteAsync(service, userData, progress, cancellationToken, args => (TResult)args.Feed);
        }

        private static Task<TResult> QueryEntryAsync<TResult>(SpreadsheetsService service, object userData, IAggregateProgress progress, CancellationToken cancellationToken)
            where TResult : AtomEntry
        {
            return ExecuteAsync(service, userData, progress, cancellationToken, args => (TResult) args.Entry);
        }

        private static Task<TResult> ExecuteAsync<TResult>(SpreadsheetsService service, object userData, IAggregateProgress progress, CancellationToken cancellationToken, Func<AsyncOperationCompletedEventArgs, TResult> getResult)
            where TResult : AtomBase
        {
            var taskSource = new TaskCompletionSource<TResult>();

            AsyncOperationProgressEventHandler progressHandler = null;
            if (progress != null)
            {
                progressHandler = (sender, eventArgs) =>
                {
                    if (eventArgs.UserState == userData)
                    {
                        progress.Report(eventArgs.ProgressPercentage);
                    }
                };

                service.AsyncOperationProgress += progressHandler;
            }

            AsyncOperationCompletedEventHandler evnetHandler = null;
            evnetHandler = (sender, eventArgs) =>
            {
                if (eventArgs.UserState == userData)
                {
                    service.AsyncOperationCompleted -= evnetHandler;
                    if (progress != null)
                    {
                        progress.Report(100);
                        service.AsyncOperationProgress -= progressHandler;
                    }

                    if (eventArgs.Error != null)
                    {
                        taskSource.TrySetException(eventArgs.Error);
                    }
                    else
                    {
                        taskSource.TrySetResult(getResult(eventArgs));
                    }
                }
            };

            service.AsyncOperationCompleted += evnetHandler;

            cancellationToken.Register(() =>
            {
                if (taskSource.TrySetCanceled())
                {
                    service.AsyncOperationCompleted -= evnetHandler;
                    if (progress != null)
                    {
                        service.AsyncOperationProgress -= progressHandler;
                    }

                    service.CancelAsync(userData);
                }
            });

            return taskSource.Task;
        }

        #endregion
    }
}
