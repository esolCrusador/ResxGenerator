using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using EnvDTE;
using ResourcesAutogenerate.DomainModels;

namespace ResxPackage.Dialog.Models
{
    public class ResourcesVm
    {
        public ResourcesVm(IEnumerable<CultureInfo> culturesList, List<string> supportedCultures, IReadOnlyCollection<Project> projects, string solutionName)
        {
            ProjectsList = new ObservableCollection<ProjectSelectItem>(projects.Select(proj => new ProjectSelectItem(proj)));
            Projects = projects;

            List<CultureSelectItem> cultures = culturesList.OrderBy(cul => cul.LCID).ThenBy(cul => cul.Name).Select(cul => new CultureSelectItem(cul, supportedCultures.Contains(cul.Name))).ToList();
            CulturesList = new ObservableCollection<CultureSelectItem>(cultures.Where(cul => !cul.IsSelected));
            SelectedCulturesList = new ObservableCollection<CultureSelectItem>(cultures.Where(cul => cul.IsSelected));

            SolutionName = solutionName;

            SyncOptions = new UpdateResourcesOptionsVm();
            
            ExternalSource = new ExternalSourceVm(Models.ExternalSource.Sync);
        }

        #region Observable

        public ObservableCollection<CultureSelectItem> CulturesList { get; set; }

        public ObservableCollection<CultureSelectItem> SelectedCulturesList { get; set; }

        public ObservableCollection<ProjectSelectItem> ProjectsList { get; set; }

        public UpdateResourcesOptionsVm SyncOptions { get; set; }

        #endregion

        public ExternalSourceVm ExternalSource { get; set; }

        public string SolutionName { get; private set; }

        public IReadOnlyCollection<Project> Projects { get; set; }

        public IReadOnlyCollection<string> SelectedCultures
        {
            get { return SelectedCulturesList.Where(cul=>cul.IsSelected).Select(c => c.CultureId).ToList(); }
        }

        public IReadOnlyCollection<Project> SelectedProjects
        {
            get { return Projects.Where(proj => ProjectsList.Where(pi => pi.IsSelected).Any(pi => pi.Equals(proj))).ToList(); }
        }

        public void UpdateSelectedCultures(bool isSelected, string cultureId)
        {
            ObservableCollection<CultureSelectItem> sourceCollection;
            ObservableCollection<CultureSelectItem> destCollection;

            if (isSelected)
            {
                sourceCollection = CulturesList;
                destCollection = SelectedCulturesList;
            }
            else
            {
                sourceCollection = SelectedCulturesList;
                destCollection = CulturesList;
            }

            CultureSelectItem cultureItem = sourceCollection.First(cul => cul.CultureId == cultureId);
            //We don't move initally selected items anywhere.
            if (cultureItem.InitiallySelected)
            {
                return;
            }
            sourceCollection.Remove(cultureItem);

            int oldIndex = destCollection.Count;
            destCollection.Add(cultureItem);

            int newIndex = destCollection
                .OrderBy(cul => cul.CultureName)
                .Select((cul, idx) => new { Culture = cul, Index = idx })
                .Where(cul => cul.Culture == cultureItem)
                .Select(cul => cul.Index)
                .First();

            destCollection.Move(oldIndex, newIndex);
        }

        /// <summary>
        /// After project processing initally selected cultures should be the same as Selected Cultures.
        /// </summary>
        public void UpdateInitiallySelectedCultures()
        {
            foreach (var cultureSelectItem in SelectedCulturesList)
            {
                cultureSelectItem.InitiallySelected = cultureSelectItem.IsSelected;
            }
        }
    }
}
