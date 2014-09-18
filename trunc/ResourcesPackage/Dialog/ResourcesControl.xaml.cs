using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Resources;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Common.Excel.Models;
using EnvDTE;
using EnvDTE100;
using GloryS.ResourcesPackage.Dialog;
using ResourcesAutogenerate;

namespace GloryS.ResourcesPackage
{
    /// <summary>
    /// Interaction logic for ResourcesControl.xaml
    /// </summary>
    public partial class ResourcesControl : UserControl
    {
        private readonly IResourceMerge _resourceMerge;

        private readonly List<string> _logMessages;

        public ResourcesControl(IResourceMerge resourceMerge, Solution dte, ILogger outputWindowLogger)
        {
            _resourceMerge = resourceMerge;
            InitializeComponent();

            _logMessages = new List<string>();
            ILogger combinedLogger = new CombinedLogger(outputWindowLogger, new DialogLogger(_logMessages));
            resourceMerge.SetLogger(combinedLogger);

            InitializeData(dte);
        }

        public ResourcesVM ViewModel { get; set; }

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

            ViewModel = new ResourcesVM(CultureInfo.GetCultures(CultureTypes.NeutralCultures), supportedCultures, projectsList);

            this.DataContext = ViewModel;
        }

        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _resourceMerge.UpdateResources(ViewModel.SelectedCultures, ViewModel.SelectedProjects, removeFiles: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            MessageBox.Show(String.Format("Resources successfully generated\r\n{0}",
                        String.Join("\r\n", _logMessages)
                        ));
            _logMessages.Clear();
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
                MessageBox.Show(ex.ToString());
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


                    MessageBox.Show(String.Format("Resources data was successfully imported.\r\n{0}",
                        String.Join("\r\n", _logMessages)
                        ));
                    _logMessages.Clear();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}