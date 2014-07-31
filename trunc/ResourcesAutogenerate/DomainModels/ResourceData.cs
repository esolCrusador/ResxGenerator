using System.Collections.Generic;
using System.Globalization;
using EnvDTE;

namespace ResourcesAutogenerate.DomainModels
{
    public class ResourceData
    {
        public ResourceData(string resourceName, string resourcePath, CultureInfo culture, ProjectItem projectItem, Dictionary<string, string> resources)
        {
            ResourceName = resourceName;
            ResourcePath = resourcePath;
            Culture = culture;
            ProjectItem = projectItem;
            Resources = resources;
        }

        public string ResourceName { get; set; }

        public string ResourcePath { get; set; }

        public CultureInfo Culture { get; set; }

        public Dictionary<string, string> Resources { get; set; }
        
        public ProjectItem ProjectItem { get; set; }
    }
}
