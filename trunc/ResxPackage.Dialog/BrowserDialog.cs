using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using ResxPackage.Resources;

namespace ResxPackage.Dialog
{
    public class BrowserDialog
    {
        private readonly Browser _browser;

        public BrowserDialog(TimeSpan timeSpan, Uri uri)
        {
            _browser = new Browser(timeSpan, uri);
        }

        public Task<DialogResult> ShowAsync(Action<string> callback)
        {
            var taskSource = new TaskCompletionSource<DialogResult>();

            _browser.CodeReceived += code =>
            {
                if (code != null)
                {
                    try
                    {
                        callback(code);
                        taskSource.TrySetResult(DialogResult.OK);
                    }
                    catch (Exception ex)
                    {
                        taskSource.TrySetException(ex);
                    }

                    _browser.Stop();
                    _browser.Hide();
                }
            };

            _browser.UrlNavigationFailed += () =>
            {
                taskSource.TrySetException(new TimeoutException(DialogRes.NavgationFailed));

                _browser.Stop();
                _browser.Hide();
            };

            _browser.Closed += (sender, args) =>
            {
                _browser.Stop();
                taskSource.TrySetResult(DialogResult.Cancel);
            };

            _browser.ShowDialog();

            return taskSource.Task;
        }
    }
}
