using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using Common.Excel.Contracts;
using EnvDTE;
using ResxPackage.Dialog.Models;
using ResxPackage.Resources;
using Window = System.Windows.Window;

namespace ResxPackage.Dialog
{
    /// <summary>
    /// Interaction logic for SelectGoogleDocument.xaml
    /// </summary>
    public partial class SelectGoogleDocument : Window
    {
        private readonly Action<string, string, DialogIcon> _showDialogAction;
        private readonly IGoogleDocumentsService _documentsService;

        public SelectGoogleDocument(Action<string, string, DialogIcon> showDialogAction, IGoogleDocumentsService documentsService)
        {
            this.Icon = PackageRes.ResxGenIcon.GetImageSource();

            _showDialogAction = showDialogAction;
            _documentsService = documentsService;

            ViewModel = new DocumentsVm();
            this.DataContext = ViewModel;

            InitializeComponent();
        }

        public DocumentsVm ViewModel { get; set; }

        public event Action<string, string> PathReceived = delegate { };

        public event Action OperationCancelled = delegate { };

        public async override void EndInit()
        {
            base.EndInit();

            await RefreshDocuments();

        }

        private async Task RefreshDocuments()
        {
            try
            {
                string selectedPath = ViewModel.GetSelectedPath();

                var documents = await _documentsService.GetDocuments();

                ViewModel.Documents.Clear();
                foreach (var document in documents)
                {
                    ViewModel.Documents.Add(new GoogleDocumentVm(document.Path, document.Url, document.Name)
                    {
                        IsSelected = selectedPath != null && document.Path == selectedPath
                    });
                }
            }
            catch (Exception e)
            {
                ShowDialogWindow(DialogIcon.Critical, DialogRes.Exception, e.ToString());

                throw;
            }
        }

        private void ShowDialogWindow(DialogIcon icon, string title, string message)
        {
            _showDialogAction(title, message, icon);
        }

        private async void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            await RefreshDocuments();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            string selectedPath = ViewModel.GetSelectedPath();
            if (selectedPath == null)
            {
                ShowDialogWindow(DialogIcon.Warning, DialogRes.DocumentNotSelected, DialogRes.PleaseSelectDocument);
                return;
            }

            string selectedUrl = ViewModel.GetSelectedUrl();

            PathReceived(selectedPath, selectedUrl);
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        public void Stop()
        {
            PathReceived = delegate { };
            OperationCancelled = delegate { };
        }

        protected override void OnClosed(EventArgs e)
        {
            OperationCancelled();
        }

        private void CreateLink_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start((string)CreateLink.Tag);
        }
    }
}
