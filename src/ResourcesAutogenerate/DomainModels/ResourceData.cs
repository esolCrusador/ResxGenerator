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
        private static readonly string StringTypeName = typeof(String).AssemblyQualifiedName;

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
            get
            {
                return _stringResources ?? (_stringResources = _resources.Where(r => r.Value.GetValueTypeName(new AssemblyName[0]) == StringTypeName)
                    .ToDictionary(r => r.Key, r => new StringResxNode(r.Value)));
            }
        }

        private IReadOnlyDictionary<string, ResXDataNode> _nptStringResources;
        public IReadOnlyDictionary<string, ResXDataNode> NotStringResources
        {
            get
            {
                return _nptStringResources ?? (_nptStringResources = _resources.Select(r => r).Where(r => r.Value.GetValueTypeName(new AssemblyName[0]) != StringTypeName)
                    .ToDictionary(r => r.Key, r => r.Value));
            }
        }

        public ProjectItem ProjectItem { get; set; }
    }
}
