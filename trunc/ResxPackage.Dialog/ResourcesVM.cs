using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using EnvDTE;

namespace ResxPackage.Dialog
{
    public class ResourcesVm
    {
        public ResourcesVm(IEnumerable<CultureInfo> culturesList, List<int> supportedCultures, IReadOnlyCollection<Project> projects, string solutionName)
        {
            ProjectsList = new ObservableCollection<ProjectSelectItem>(projects.Select(proj => new ProjectSelectItem(proj)));
            Projects = projects;
            CulturesList = new ObservableCollection<CultureSelectItem>(culturesList.Select(cul => new CultureSelectItem(cul, supportedCultures.Contains(cul.LCID))).OrderBy(cul => cul.CultureName));
            SolutionName = solutionName;
        }

        #region Observable

        public ObservableCollection<CultureSelectItem> CulturesList { get; set; }

        public ObservableCollection<ProjectSelectItem> ProjectsList { get; set; }

        #endregion

        public string SolutionName { get; private set; }

        public IReadOnlyCollection<Project> Projects { get; set; }

        public IReadOnlyCollection<int> SelectedCultures
        {
            get { return CulturesList.Where(c => c.IsSelected).Select(c => c.CultureId).ToList(); }
        }

        public IReadOnlyCollection<Project> SelectedProjects
        {
            get { return Projects.Where(proj => ProjectsList.Where(pi => pi.IsSelected).Any(pi => pi.Equals(proj))).ToList(); }
        }
    }
}
