using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Common.Excel.Contracts;

namespace ResxPackage.Dialog
{
    public class SelectGoogleDocumentDialog
    {
        private readonly SelectGoogleDocument _window;

        public SelectGoogleDocumentDialog(Action<string, string, DialogIcon> showDialogAction, IGoogleDocumentsService documentsService)
        {
            _window = new SelectGoogleDocument(showDialogAction, documentsService);
        }

        public Task<DialogResult> ShowAsync(Action<string, string> callback)
        {
            TaskCompletionSource<DialogResult> completionSource = new TaskCompletionSource<DialogResult>();

            _window.PathReceived += (path, url) =>
            {
                try
                {
                    callback(path, url);

                    completionSource.TrySetResult(DialogResult.OK);
                }
                catch (Exception e)
                {
                    completionSource.TrySetException(e);
                }
                _window.Stop();
                _window.Hide();
            };

            _window.OperationCancelled += () =>
            {
                completionSource.TrySetResult(DialogResult.Cancel);

                _window.Stop();
            };

            _window.ShowDialog();

            return completionSource.Task;
        }
    }
}
