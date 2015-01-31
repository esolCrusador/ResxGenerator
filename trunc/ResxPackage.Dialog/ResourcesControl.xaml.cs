using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;
using Common.Excel.Implementation;
using EnvDTE;
using Google.GData.Client;
using Google.GData.Spreadsheets;
using ResourcesAutogenerate;
using ResourcesAutogenerate.DomainModels;
using ResxPackage.Dialog;
using ResxPackage.Dialog.Models;
using ResxPackage.Resources;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Process = System.Diagnostics.Process;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;
using TextBox = System.Windows.Controls.TextBox;
using UserControl = System.Windows.Controls.UserControl;

namespace GloryS.ResourcesPackage
{
    /// <summary>
    /// Interaction logic for ResourcesControl.xaml
    /// </summary>
    public partial class ResourcesControl : UserControl
    {
        private readonly IResourceMerge _resourceMerge;
        private readonly Action<string, string, DialogIcon> _showDialogAction;
        private readonly List<string> _logMessages;

        private readonly ExcelGenerator _excelGenerator;

        private OAuth2Parameters _googleOAuth2Parameters;
        private GoogleDocGenerator _googleDocGenerator;

        private StatusProgress _statusProgress;
        private CancellationTokenSource _cancellationTokenSource;

        public ResourcesControl(IResourceMerge resourceMerge, Solution dte, ILogger outputWindowLogger, Action<string, string, DialogIcon> showDialogAction)
        {
            InitializeComponent();

            _excelGenerator = new ExcelGenerator();
            _resourceMerge = resourceMerge;
            _showDialogAction = showDialogAction;
            _logMessages = new List<string>();
            ILogger combinedLogger = new CombinedLogger(outputWindowLogger, new DialogLogger(_logMessages));
            resourceMerge.SetLogger(combinedLogger);
            _statusProgress = new StatusProgress(
                p =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        StatusProgress.Value = p;
                        StatusProgressText.Text = (int)Math.Round(p) + "%";
                    });
                },
                (s, p) =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        StatusProgress.Value = p;
                        StatusProgressText.Text = (int)Math.Round(p) + "%";
                        StatusText.Text = s;
                    });
                }
                );
            _cancellationTokenSource = new CancellationTokenSource();

            InitializeData(dte);
        }

        public ResourcesVm ViewModel { get; set; }

        public override void EndInit()
        {
            base.EndInit();

            this.GenResxIcon.Source = DialogRes.ResxGen.GetImageSource();

            this.ExportToExcelIcon.Source = DialogRes.ExportToExcel.GetImageSource();
            this.ImportFromExcelIcon.Source = DialogRes.ImpotrFromExcel.GetImageSource();
            this.GenResxSmallIcon.Source = DialogRes.ResxGen.GetImageSource();

            this.ExportToGSheetsIcon.Source = DialogRes.ExportToGSheets.GetImageSource();
            this.ImportFromGSheetsIcon.Source = DialogRes.ImpotrFromGSheets.GetImageSource();

            this.ExcelIcon.Source = DialogRes.Excel.GetImageSource();
            this.GoogleSheetsIcon.Source = DialogRes.GSheets.GetImageSource();
        }

        private void InitializeData(Solution solution)
        {
            List<Project> projectsList = solution.GetAllProjects()
                .Where(proj =>
                    proj.GetAllItems().Any(item => Path.GetExtension(item.Name) == ".resx")
                )
                .ToList();

            List<int> supportedCultures = projectsList
                .SelectMany(proj => proj.GetAllItems().Select(item => item.Name).Where(itemName => Path.GetExtension(itemName) == ".resx")
                    .Select(Path.GetFileNameWithoutExtension)
                    .Select(fileName =>
                    {
                        string cultureName = Path.GetExtension(fileName);

                        int culture;
                        if (String.IsNullOrEmpty(cultureName))
                        {
                            culture =  CultureInfo.InvariantCulture.LCID;
                        }
                        else
                        {
                            try
                            {
                                culture = CultureInfo.GetCultureInfo(cultureName.Substring(1)).LCID;
                            }
                            catch (CultureNotFoundException)
                            {
                                culture = CultureInfo.InvariantCulture.LCID;
                            }
                        }

                        return culture;
                    }))
                .Distinct()
                .ToList();

            ViewModel = new ResourcesVm(CultureInfo.GetCultures(CultureTypes.NeutralCultures), supportedCultures, projectsList, Path.GetFileNameWithoutExtension(solution.FullName));

            this.DataContext = ViewModel;
        }

        private void ShowDialogWindow(DialogIcon icon, string title, string message)
        {
            _showDialogAction(title, message, icon);
        }

        private async void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _statusProgress.Report(StatusRes.GeneratingResx);
                ShowOverlay(GenResxIcon);
                await _resourceMerge.UpdateResourcesAsync(ViewModel.SelectedCultures, ViewModel.SelectedProjects, _statusProgress, _cancellationTokenSource.Token, ViewModel.SyncOptions.GetOptions());
                ViewModel.UpdateInitiallySelectedCultures();

                HideOverlay();
                ShowDialogWindow(DialogIcon.Info, DialogRes.Success, String.Format(LoggerRes.SuccessfullyGeneratedFormat, String.Join(LoggerRes.Delimiter, _logMessages)));
            }
            catch (OperationCanceledException)
            {
                HideOverlay();
                ShowDialogWindow(DialogIcon.Info, DialogRes.Success, LoggerRes.OperationCancelled);
            }
            catch (Exception ex)
            {
                HideOverlay();
                ShowDialogWindow(DialogIcon.Critical, DialogRes.Exception, ex.ToString() + LoggerRes.Delimiter + String.Join(LoggerRes.Delimiter, _logMessages));

                throw;
            }
            finally
            {
                _statusProgress.Clear();
                _logMessages.Clear();
            }
        }

        private async void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveFileDialog = new SaveFileDialog();
                saveFileDialog.AddExtension = true;
                saveFileDialog.FileName = "Resources";
                saveFileDialog.DefaultExt = ".xlsx";
                saveFileDialog.Filter = "Excel document (.xlsx)|*.xlsx";


                if (saveFileDialog.ShowDialog() == true)
                {
                    ShowOverlay(GenResxIcon);
                    await _resourceMerge.UpdateResourcesAsync(ViewModel.SelectedCultures, ViewModel.SelectedProjects, _statusProgress, _cancellationTokenSource.Token, new UpdateResourcesOptions{RemoveNotSelectedCultures = false});
                    ViewModel.UpdateInitiallySelectedCultures();

                    ShowOverlay(ExportToExcelIcon);
                    await _resourceMerge.ExportToDocumentAsync(_excelGenerator, saveFileDialog.FileName, ViewModel.SelectedCultures, ViewModel.SelectedProjects, _statusProgress, _cancellationTokenSource.Token);

                    HideOverlay();
                    Process.Start("explorer.exe", String.Format("/n /select,{0},{1}", Path.GetDirectoryName(saveFileDialog.FileName), Path.GetFileName(saveFileDialog.FileName)));
                }
            }
            catch (OperationCanceledException)
            {
                HideOverlay();
                ShowDialogWindow(DialogIcon.Info, DialogRes.Success, LoggerRes.OperationCancelled);
            }
            catch (Exception ex)
            {
                HideOverlay();
                ShowDialogWindow(DialogIcon.Critical, DialogRes.Exception, ex.ToString());
            }
            finally
            {
                _logMessages.Clear();
            }
        }

        private async void ImportFromExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openFileDialog = new OpenFileDialog();
                openFileDialog.AddExtension = true;
                openFileDialog.FileName = "Resources";
                openFileDialog.DefaultExt = ".xlsx";
                openFileDialog.Filter = "Excel document (.xlsx)|*.xlsx";


                if (openFileDialog.ShowDialog() == true)
                {
                    ShowOverlay(GenResxIcon);
                    await _resourceMerge.UpdateResourcesAsync(ViewModel.SelectedCultures, ViewModel.SelectedProjects, _statusProgress, _cancellationTokenSource.Token, new UpdateResourcesOptions{RemoveNotSelectedCultures = false});
                    using (var reader = File.OpenRead(openFileDialog.FileName))
                    {
                        byte[] buffer = new byte[reader.Length];

                        reader.Read(buffer, 0, (int) reader.Length);

                        ShowOverlay(ImportFromExcelIcon);
                        await _resourceMerge.ImportFromDocumentAsync(_excelGenerator, openFileDialog.FileName, ViewModel.SelectedCultures, ViewModel.SelectedProjects, _statusProgress, _cancellationTokenSource.Token);
                    }

                    HideOverlay();
                    ShowDialogWindow(DialogIcon.Info, DialogRes.Success, String.Format(LoggerRes.SuccessfullyImportedFormat, String.Join(LoggerRes.Delimiter, _logMessages)));
                }
            }
            catch (OperationCanceledException)
            {
                HideOverlay();
                ShowDialogWindow(DialogIcon.Info, DialogRes.Success, LoggerRes.OperationCancelled);
            }
            catch (Exception ex)
            {
                HideOverlay();
                ShowDialogWindow(DialogIcon.Critical, DialogRes.Exception, ex.ToString() + LoggerRes.Delimiter + String.Join(LoggerRes.Delimiter, _logMessages));

                throw;
            }
            finally
            {
                _logMessages.Clear();
            }
        }

        private void CheckBoxZone_CheckChanged(object sender, RoutedEventArgs e)
        {
            var checkBox = (ToggleButton)sender;
            var cultureId = (int)checkBox.Tag;

            ViewModel.UpdateSelectedCultures(checkBox.IsChecked.HasValue && checkBox.IsChecked.Value, cultureId);
        }

        private void FilterCultures_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox filterBox = (TextBox) sender;
            string text = filterBox.Text;

            foreach (CultureSelectItem cultureItem in ViewModel.CulturesList)
            {
                cultureItem.IsVisible = cultureItem.CultureName.IndexOf(text, StringComparison.OrdinalIgnoreCase) != -1;
            }
        }

        private void ShowOverlay(Image image)
        {
            OverlayImage.Source = image.Source;
            Overlay.Visibility = Visibility.Visible;
        }

        public void HideOverlay()
        {
            OverlayImage.Source = null;
            Overlay.Visibility = Visibility.Hidden;
        }

        private async Task<DialogResult> GetGoogleService(Action<string, string> setPath)
        {
            if (_googleDocGenerator == null)
            {
                OAuth2Parameters parameters = new OAuth2Parameters
                {
                    ClientId = "261863669828-1k61kiqfjcci0psjr5e00vcpnsnllnug.apps.googleusercontent.com",
                    ClientSecret = "IDtucbpfYi3C7zWxsJUX4HbV",
                    RedirectUri = "urn:ietf:wg:oauth:2.0:oob",
                    Scope = "https://spreadsheets.google.com/feeds https://docs.google.com/feeds",
                    ResponseType = "code"
                };
                string authorizationUrl = OAuthUtil.CreateOAuth2AuthorizationUrl(parameters);

                var browserDialog = new BrowserDialog(TimeSpan.FromSeconds(30), new Uri(authorizationUrl));

                try
                {
                    if (await browserDialog.ShowAsync(code => parameters.AccessCode = code) != DialogResult.OK)
                    {
                        return DialogResult.Cancel;
                    }
                }
                catch (Exception e)
                {
                    ShowDialogWindow(DialogIcon.Critical, DialogRes.Exception, e.ToString());

                    throw;
                }

                OAuthUtil.GetAccessToken(parameters);

                _googleOAuth2Parameters = parameters;

                var service = new SpreadsheetsService("MySpreadsheetIntegration-v1")
                {
                    RequestFactory = new GOAuth2RequestFactory(null, "MySpreadsheetIntegration-v1", parameters)
                };

                _googleDocGenerator = new GoogleDocGenerator(service);
            }
            else
            {
                if (_googleOAuth2Parameters.TokenExpiry <= DateTime.Now)
                {
                    OAuthUtil.RefreshAccessToken(_googleOAuth2Parameters);
                }
            }

            SelectGoogleDocumentDialog selectDocumentDialog = new SelectGoogleDocumentDialog(_showDialogAction, _googleDocGenerator);
            
            DialogResult dialogResult = await selectDocumentDialog.ShowAsync(setPath);

            return dialogResult;
        }

        private async void ExportToGSheets_Click(object sender, RoutedEventArgs e)
        {
            string documentPath = null;
            string documentPublicUrl = null;

            try
            {
                if (await GetGoogleService((path, publicUrl) => { documentPath = path; documentPublicUrl = publicUrl; }) == DialogResult.OK)
                {
                    _statusProgress.Report(StatusRes.GeneratingResx);
                    ShowOverlay(GenResxIcon);
                    await _resourceMerge.UpdateResourcesAsync(ViewModel.SelectedCultures, ViewModel.SelectedProjects, _statusProgress, _cancellationTokenSource.Token, new UpdateResourcesOptions{RemoveNotSelectedCultures = false});
                    ViewModel.UpdateInitiallySelectedCultures();

                    _statusProgress.Report(StatusRes.ExportToGSheets);
                    ShowOverlay(ExportToGSheetsIcon);
                    await _resourceMerge.ExportToDocumentAsync(_googleDocGenerator, documentPath, ViewModel.SelectedCultures, ViewModel.SelectedProjects, _statusProgress, _cancellationTokenSource.Token);

                    HideOverlay();
                    Process.Start(documentPublicUrl);
                }
            }
            catch (OperationCanceledException)
            {
                HideOverlay();
                ShowDialogWindow(DialogIcon.Info, DialogRes.Success, LoggerRes.OperationCancelled);
            }
            catch (Exception ex)
            {
                HideOverlay();
                ShowDialogWindow(DialogIcon.Critical, DialogRes.Exception, ex.ToString());

                throw;
            }
        }

        private void CancelOperation_Click(object sender, RoutedEventArgs e)
        {
            _cancellationTokenSource.Cancel();
            _cancellationTokenSource = new CancellationTokenSource();
        }

        private async void ImportFromGSheets_Click(object sender, RoutedEventArgs e)
        {
            string documentPath = null;

            try
            {
                if (await GetGoogleService((path, publicUrl) =>
                {
                    documentPath = path;
                }) == DialogResult.OK)
                {
                    _statusProgress.Report(StatusRes.GeneratingResx);
                    ShowOverlay(GenResxIcon);
                    await _resourceMerge.UpdateResourcesAsync(ViewModel.SelectedCultures, ViewModel.SelectedProjects, _statusProgress, _cancellationTokenSource.Token, new UpdateResourcesOptions{RemoveNotSelectedCultures = false});
                    ViewModel.UpdateInitiallySelectedCultures();

                    _statusProgress.Report(StatusRes.ImportFromGSheets);
                    ShowOverlay(ImportFromGSheetsIcon);
                    await _resourceMerge.ImportFromDocumentAsync(_googleDocGenerator, documentPath, ViewModel.SelectedCultures, ViewModel.SelectedProjects, _statusProgress, _cancellationTokenSource.Token);

                    HideOverlay();
                    ShowDialogWindow(DialogIcon.Info, DialogRes.Success, String.Format(LoggerRes.SuccessfullyImportedFormat, String.Join(LoggerRes.Delimiter, _logMessages)));
                }
            }
            catch (OperationCanceledException)
            {
                HideOverlay();
                ShowDialogWindow(DialogIcon.Info, DialogRes.Success, LoggerRes.OperationCancelled);
            }
            catch (Exception ex)
            {
                HideOverlay();
                ShowDialogWindow(DialogIcon.Critical, DialogRes.Exception, ex.ToString());

                throw;
            }
            finally
            {
                _logMessages.Clear();
            }
        }
    }
} ;