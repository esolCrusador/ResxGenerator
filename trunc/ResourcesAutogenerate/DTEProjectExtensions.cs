using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;

namespace ResourcesAutogenerate
{
    public static class DTEProjectExtensions
    {
        public static IEnumerable<ProjectItem> GetAllItems(this Project project)
        {
            return GetProjectItems(project.ProjectItems.Cast<ProjectItem>());
        }

        private static IEnumerable<ProjectItem> GetProjectItems(IEnumerable<ProjectItem> projectItems)
        {
            var projectsList = projectItems as IList<ProjectItem> ?? projectItems.ToList();

            return projectsList.Concat(projectsList.SelectMany(pi => GetProjectItems(pi.ProjectItems == null ? Enumerable.Empty<ProjectItem>() : pi.ProjectItems.Cast<ProjectItem>())));
        }

        public static IEnumerable<Project> GetAllProjects(this Solution solution)
        {
            return solution.Projects.Cast<Project>().SelectMany(proj => proj.GetProjectWithSubProjects());

        }

        private static IEnumerable<Project> GetProjectWithSubProjects(this Project project)
        {
            return Enumerable.Repeat(project, 1).Concat(
                project.GetAllItems()
                    .Where(item => item.SubProject != null)
                    .SelectMany(item => GetProjectWithSubProjects(item.SubProject))
                );
        }
    }
}
