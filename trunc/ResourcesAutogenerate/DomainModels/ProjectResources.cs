using System.Collections.Generic;
using EnvDTE;

namespace ResourcesAutogenerate.DomainModels
{
    public class ProjectResources
    {
        public string ProjectId { get; set; }

        public string ProjectName { get; set; }

        public string ProjectDirectory { get; set; }

        public IReadOnlyList<ProjectItem> ResourceProjectItems { get; set; }

        public Dictionary<string, Dictionary<int, ResourceData>> Resources { get; set; }
    }
}
