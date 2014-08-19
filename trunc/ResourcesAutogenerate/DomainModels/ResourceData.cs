using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Resources;
using EnvDTE;

namespace ResourcesAutogenerate.DomainModels
{
    public class ResourceData
    {
        public ResourceData(string resourceName, string resourcePath, CultureInfo culture, ProjectItem projectItem, Dictionary<string, ResXDataNode> resources)
        {
            ResourceName = resourceName;
            ResourcePath = resourcePath;
            Culture = culture;
            ProjectItem = projectItem;
            _resources = resources;
        }

        public string ResourceName { get; set; }

        public string ResourcePath { get; set; }

        public CultureInfo Culture { get; set; }

        private readonly IReadOnlyDictionary<string, ResXDataNode> _resources;
        public IReadOnlyDictionary<string, ResXDataNode> Resources
        {
            get { return _resources; }
        }

        private IReadOnlyDictionary<string, StringResxNode> _stringResources;
        public IReadOnlyDictionary<string, StringResxNode> StringResources
        {
            get { return _stringResources ?? (_stringResources = _resources.Where(r => r.Value.FileRef == null).ToDictionary(r => r.Key, r => new StringResxNode(r.Value))); }
        }

        private IReadOnlyDictionary<string, ResXDataNode> _fileResources;
        public IReadOnlyDictionary<string, ResXDataNode> FileResources
        {
            get { return _fileResources ?? (_fileResources = _resources.Select(r => r).Where(r => r.Value.FileRef != null).ToDictionary(r => r.Key, r => r.Value)); }
        }

        public ProjectItem ProjectItem { get; set; }
    }
}
