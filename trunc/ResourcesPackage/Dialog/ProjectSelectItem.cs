using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;

namespace GloryS.ResourcesPackage
{
    public class ProjectSelectItem: IEquatable<EnvDTE.Project>
    {
        public ProjectSelectItem()
        {
            
        }

        public ProjectSelectItem(Project project, bool isSelected)
        {
            ProjectId = project.UniqueName;
            ProjectName = project.Name;

            IsSelected = isSelected;
        }

        public ProjectSelectItem(Project project)
            :this(project, true)
        {
            
        }

        public string ProjectId { get; set; }

        public string ProjectName { get; set; }

        public bool IsSelected { get; set; }
        
        public bool Equals(Project other)
        {
            return ProjectId == other.UniqueName;
        }
    }
}
