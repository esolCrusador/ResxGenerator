using System;
using System.Threading;
using System.Windows;
using System.Windows.Navigation;
using mshtml;
using ResxPackage.Resources;

namespace ResxPackage.Dialog
{
    /// <summary>
    /// Interaction logic for Browser.xaml
    /// </summary>
    public partial class Browser : Window
    {
        private readonly TimeSpan _timeout;
        private Uri _url;
        private readonly Timer _timer;

        public Browser(TimeSpan timeout, Uri uri)
        {
            this.Icon = PackageRes.ResxGenIcon.GetImageSource();

            _timeout = timeout;
            InitializeComponent();

            _timer = new Timer(OnTimeoutTimer, null, -1, -1);
            Url = uri;
        }
        
        public event Action<string> CodeReceived = delegate { };
        public event Action NavigationCancelled = delegate { };
        public event Action UrlNavigationFailed = delegate { };

        public Uri Url
        {
            get { return _url; }
            set
            {
                _url = value;

                DelayText.Visibility = Visibility.Collapsed;
                WebBrowser.Visibility = Visibility.Collapsed;            
                WebBrowser.Navigate(value);
            }
        }

        public void Stop()
        {
            _timer.Change(-1, -1);

            CodeReceived = delegate { };
            UrlNavigationFailed = delegate { };
        }

        public void SetDelayMode()
        {
            WebBrowser.Visibility = Visibility.Collapsed;
            DelayText.Visibility = Visibility.Visible;
        }

        public void Reload()
        {
            Url = Url;
        }

        private void WebBrowser_OnNavigating(object sender, NavigatingCancelEventArgs navigatingCancelEventArgs)
        {
            WebBrowser.IsEnabled = false;

            _timer.Change(_timeout, _timeout);
        }

        private void OnTimeoutTimer(object state)
        {
            _timer.Change(-1, -1);

            UrlNavigationFailed();
        }

        private void WebBrowser_LoadCompleted(object sender, NavigationEventArgs e)
        {
            _timer.Change(-1, -1);

            WebBrowser.IsEnabled = true;

            WebBrowser.Visibility = Visibility.Visible;

            if (e.Uri.AbsolutePath == "/o/oauth2/approval")
            {
                HTMLDocument document = ((HTMLDocument)WebBrowser.Document);
                IHTMLElement element = document.getElementById("code");
                if (element == null)
                {
                    Close();
                }
                else
                {
                    IHTMLInputTextElement input = (IHTMLInputTextElement) element;

                    string code = input.value;

                    CodeReceived(code);
                }
            }
        }
    }
}
