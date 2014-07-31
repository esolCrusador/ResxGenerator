using System.Collections.Generic;
using Common.Excel.Models;
using EnvDTE;

namespace ResourcesAutogenerate
{
    public interface IResourceMerge
    {
        void UpdateResources(IReadOnlyCollection<int> selectedCultures, IReadOnlyCollection<Project> selectedProjects, bool removeFiles = true);

        FileInfoContainer ExportToExcelFile(IReadOnlyCollection<int> selectedCultures, IReadOnlyCollection<Project> selectedProjects, string title);
        void ImportFromExcel(IReadOnlyCollection<int> selectedCultures, IReadOnlyCollection<Project> selectedProjects, FileInfoContainer file);
    }
}
