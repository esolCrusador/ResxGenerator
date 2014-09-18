using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using Common.Excel.Models;
using EnvDTE;
using ResourcesAutogenerate;
using ResxPackage.Dialog;
using ResxPackage.Resources;

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

        public ResourcesControl(IResourceMerge resourceMerge, Solution dte, ILogger outputWindowLogger, Action<string, string, DialogIcon> showDialogAction)
        {
            InitializeComponent();

            _resourceMerge = resourceMerge;
            _showDialogAction = showDialogAction;
            _logMessages = new List<string>();
            ILogger combinedLogger = new CombinedLogger(outputWindowLogger, new DialogLogger(_logMessages));
            resourceMerge.SetLogger(combinedLogger);

            InitializeData(dte);
        }

        public ResourcesVm ViewModel { get; set; }

        public override void EndInit()
        {
            base.EndInit();

            this.GenResxIcon.Source = DialogRes.ResxGen.GetImageSource();
            this.ExportToExcelIcon.Source = DialogRes.ExportToExcel.GetImageSource();
            this.ImportFromExcelIcon.Source = DialogRes.ImpotrFromExcel.GetImageSource();
        }

        private void InitializeData(Solution solution)
        {
            List<Project> projectsList = solution.GetAllProjects()
                .Where(proj =>
                    proj.GetAllItems().Any(item => System.IO.Path.GetExtension(item.Name) == ".resx")
                )
                .ToList();

            List<int> supportedCultures = projectsList
                .SelectMany(proj => proj.GetAllItems().Select(item => item.Name).Where(itemName => System.IO.Path.GetExtension(itemName) == ".resx")
                    .Select(System.IO.Path.GetFileNameWithoutExtension)
                    .Select(fileName =>
                    {
                        string culture = System.IO.Path.GetExtension(fileName);

                        return String.IsNullOrEmpty(culture) ? CultureInfo.InvariantCulture.LCID : CultureInfo.GetCultureInfo(culture.Substring(1)).LCID;
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

        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _resourceMerge.UpdateResources(ViewModel.SelectedCultures, ViewModel.SelectedProjects, removeFiles: true);
                ViewModel.UpdateInitiallySelectedCultures();

                ShowDialogWindow(DialogIcon.Info, DialogRes.Success, String.Format(LoggerRes.SuccessfullyGeneratedFormat, String.Join(LoggerRes.Delimiter, _logMessages)));
            }
            catch (Exception ex)
            {
                ShowDialogWindow(DialogIcon.Critical, DialogRes.Exception, ex.ToString() + LoggerRes.Delimiter + String.Join(LoggerRes.Delimiter, _logMessages));

                throw;
            }
            finally
            {
                _logMessages.Clear();
            }
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveFileDialog = new Microsoft.Win32.SaveFileDialog();
                saveFileDialog.AddExtension = true;
                saveFileDialog.FileName = "Resources";
                saveFileDialog.DefaultExt = ".xlsx";
                saveFileDialog.Filter = "Excel document (.xlsx)|*.xlsx";


                if (saveFileDialog.ShowDialog() == true)
                {
                    _resourceMerge.UpdateResources(ViewModel.SelectedCultures, ViewModel.SelectedProjects, removeFiles: false);
                    ViewModel.UpdateInitiallySelectedCultures();

                    var result = _resourceMerge.ExportToExcelFile(ViewModel.SelectedCultures, ViewModel.SelectedProjects, System.IO.Path.GetFileNameWithoutExtension(saveFileDialog.FileName));

                    if (File.Exists(saveFileDialog.FileName))
                    {
                        File.Delete(saveFileDialog.FileName);
                    }

                    using (FileStream fileStream = File.Create(saveFileDialog.FileName))
                    {
                        fileStream.Write(result.Bytes, 0, result.Bytes.Length);
                    }

                    System.Diagnostics.Process.Start("explorer.exe", String.Format("/n /select,{0},{1}", System.IO.Path.GetDirectoryName(saveFileDialog.FileName), System.IO.Path.GetFileName(saveFileDialog.FileName)));
                }
            }
            catch (Exception ex)
            {
                ShowDialogWindow(DialogIcon.Critical, DialogRes.Exception, ex.ToString());
            }
        }

        private void ImportFromExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openFileDialog = new Microsoft.Win32.OpenFileDialog();
                openFileDialog.AddExtension = true;
                openFileDialog.FileName = "Resources";
                openFileDialog.DefaultExt = ".xlsx";
                openFileDialog.Filter = "Excel document (.xlsx)|*.xlsx";


                if (openFileDialog.ShowDialog() == true)
                {
                    _resourceMerge.UpdateResources(ViewModel.SelectedCultures, ViewModel.SelectedProjects, removeFiles: false);
                    using (var reader = File.OpenRead(openFileDialog.FileName))
                    {
                        byte[] buffer = new byte[reader.Length];

                        reader.Read(buffer, 0, (int) reader.Length);

                        _resourceMerge.ImportFromExcel(ViewModel.SelectedCultures, ViewModel.SelectedProjects, new FileInfoContainer(buffer, openFileDialog.FileName));
                    }

                    ShowDialogWindow(DialogIcon.Info, DialogRes.Success, String.Format(LoggerRes.SuccessfullyImportedFormat, String.Join(LoggerRes.Delimiter, _logMessages)));
                }
            }
            catch (Exception ex)
            {
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
    }
} ;