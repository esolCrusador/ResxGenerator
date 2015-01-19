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

        public static Task<SpreadsheetFeed> GetFeedAsync(this SpreadsheetsService service, SpreadsheetQuery query, object userData, CancellationToken cancellationToken = default(CancellationToken))
        {
            return GetFeedAsync<SpreadsheetFeed, SpreadsheetQuery>(service, query, userData, cancellationToken);
        }

        public static Task<WorksheetFeed> GetFeedAsync(this SpreadsheetsService service, WorksheetQuery query, object userData, CancellationToken cancellationToken = default(CancellationToken))
        {
            return GetFeedAsync<WorksheetFeed, WorksheetQuery>(service, query, userData, cancellationToken);
        }

        public static Task<CellFeed> GetFeedAsync(this SpreadsheetsService service, CellQuery query, object userData, CancellationToken cancellationToken = default(CancellationToken))
        {
            return GetFeedAsync<CellFeed, CellQuery>(service, query, userData, cancellationToken);
        }

        public static Task<ListFeed> GetFeedAsync(this SpreadsheetsService service, ListQuery query, object userData, CancellationToken cancellationToken = default(CancellationToken))
        {
            return GetFeedAsync<ListFeed, ListQuery>(service, query, userData, cancellationToken);
        }

        private static Task<TResult> GetFeedAsync<TResult, TQuery>(SpreadsheetsService service, TQuery query, object userData, CancellationToken cancellationToken)
            where TQuery : FeedQuery
            where TResult : AtomFeed
        {
            service.QueryFeedAync(query.Uri, query.ModifiedSince, userData);

            return QueryFeedAsync<TResult>(service, userData, cancellationToken);
        }

        #endregion 

        #region Async Batch

        public static Task<TFeed> BatchFeedAsync<TFeed>(this SpreadsheetsService service, TFeed feed, object userData, CancellationToken cancellationToken = default (CancellationToken))
            where TFeed: AtomFeed
        {
            service.BatchAsync(feed, new Uri(feed.Batch), userData);

            return QueryFeedAsync<TFeed>(service, userData, cancellationToken);
        }

        #endregion

        #region Async Entry Operations

        public static Task DeleteItemAsync<TEntry>(this SpreadsheetsService service, TEntry item, object userData, CancellationToken cancellationToken = default (CancellationToken))
            where  TEntry: AtomEntry
        {
            service.DeleteAsync(item, true, userData);

            return QueryEntryAsync<TEntry>(service, userData, cancellationToken);
        }

        public static Task<TEntry> InsertItemAsync<TEntry>(this SpreadsheetsService service, string feedUri, TEntry item, object userData, CancellationToken cancellationToken = default (CancellationToken))
            where TEntry: AtomEntry
        {
            service.InsertAsync(new Uri(feedUri), item, userData);

            return QueryEntryAsync<TEntry>(service, userData, cancellationToken);
        }

        public static Task<TEntry> UpdateItemAsync<TEntry>(this SpreadsheetsService service, TEntry item, object userData, CancellationToken cancellationToken = default (CancellationToken))
            where TEntry : AtomEntry
        {
            service.UpdateAsync(item, userData);

            return QueryEntryAsync<TEntry>(service, userData, cancellationToken);
        }

        #endregion

        #region Event To Task conversion

        private static Task<TResult> QueryFeedAsync<TResult>(SpreadsheetsService service, object userData, CancellationToken cancellationToken)
            where TResult: AtomFeed
        {
            return ExecuteAsync(service, userData, cancellationToken, args => (TResult)args.Feed);
        }

        private static Task<TResult> QueryEntryAsync<TResult>(SpreadsheetsService service, object userData, CancellationToken cancellationToken)
            where TResult : AtomEntry
        {
            return ExecuteAsync(service, userData, cancellationToken, args => (TResult) args.Entry);
        }

        private static Task<TResult> ExecuteAsync<TResult>(SpreadsheetsService service, object userData, CancellationToken cancellationToken, Func<AsyncOperationCompletedEventArgs, TResult> getResult)
            where TResult : AtomBase
        {
            var taskSource = new TaskCompletionSource<TResult>();

            AsyncOperationCompletedEventHandler evnetHandler = null;
            evnetHandler = (sender, eventArgs) =>
            {
                if (eventArgs.UserState == userData)
                {
                    service.AsyncOperationCompleted -= evnetHandler;
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
                    service.CancelAsync(userData);
                }
            });

            return taskSource.Task;
        }

        #endregion
    }
}
