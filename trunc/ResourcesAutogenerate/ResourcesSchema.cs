using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Resources;
using Common.Excel.Contracts;
using Common.Excel.Models;
using EnvDTE;
using ResourcesAutogenerate.DomainModels;

namespace ResourcesAutogenerate
{
    public class ResourcesSchema : IResourceMerge
    {
        private const string InvariantCultureDisplayName = "Default";
        private static readonly int InvariantCultureId = CultureInfo.InvariantCulture.LCID;

        private readonly IExcelGenerator _excelGenerator;

        public ResourcesSchema(IExcelGenerator excelGenerator)
        {
            _excelGenerator = excelGenerator;
        }

        public void UpdateResources(IReadOnlyCollection<int> selectedCultures, IReadOnlyCollection<Project> selectedProjects, bool removeFiles = true)
        {
            IReadOnlyDictionary<int, CultureInfo> selectedCultureInfos = selectedCultures.Select(CultureInfo.GetCultureInfo)
                .ToDictionary(cult => cult.LCID, cult => cult);

            IReadOnlyDictionary<string, Project> projectsDictionary = selectedProjects.ToDictionary(proj => proj.UniqueName, proj => proj);

            SolutionResources solutionResources = GetSolutionResources(null, selectedProjects);

            foreach (var projectResources in solutionResources.ProjectResources)
            {
                var project = projectsDictionary[projectResources.ProjectId];

                foreach (Dictionary<int, ResourceData> resourceFileGroup in projectResources.Resources.Values)
                {
                    var projectItems2Remove = resourceFileGroup.Where(f => !selectedCultureInfos.ContainsKey(f.Key))
                        .Join(project.ProjectItems.Cast<ProjectItem>(), f => f.Value.ResourcePath, item => item.FileNames[0], (f, item) => new {ProjectItem = item, CultureId = f.Key})
                        .ToList();

                    var cultures2Add = selectedCultureInfos.Where(cult => !resourceFileGroup.ContainsKey(cult.Key)).Select(cult => cult.Value).ToList();

                    if (removeFiles)
                    {
                        foreach (var projectItem in projectItems2Remove)
                        {
                            projectItem.ProjectItem.Delete();
                            resourceFileGroup.Remove(projectItem.CultureId);
                        }
                    }

                    var neutralCulture = resourceFileGroup[InvariantCultureId];

                    foreach (var cultureInfo in cultures2Add)
                    {
                        string resourcePath = Path.Combine(Path.GetDirectoryName(neutralCulture.ResourcePath), neutralCulture.ResourceName + "." + cultureInfo.TwoLetterISOLanguageName.ToUpper() + ".resx");

                        using (File.Create(resourcePath)) { }

                        ProjectItem projectItem = project.ProjectItems.AddFromFile(resourcePath);

                        var newFile = new ResourceData
                            (
                            resourceName: neutralCulture.ResourceName,
                            resourcePath: resourcePath,
                            culture: cultureInfo,
                            projectItem: projectItem,
                            resources: new Dictionary<string, string>(0)
                            );

                        resourceFileGroup.Add(cultureInfo.LCID, newFile);
                    }

                    var otherCultureResources = resourceFileGroup.Where(resData => resData.Key != InvariantCultureId).Select(resData => resData.Value).ToList();

                    UpdateHierarchy(neutralCulture.ProjectItem, otherCultureResources.Select(val => val.ProjectItem).ToList());
                    UpdateResourceFiles(neutralCulture, otherCultureResources);

                    project.Save();
                }
            }
        }

        public FileInfoContainer ExportToExcelFile(IReadOnlyCollection<int> selectedCultures, IReadOnlyCollection<Project> selectedProjects, string title)
        {
            SolutionResources solutionResources = GetSolutionResources(selectedCultures, selectedProjects);
            


            var cultures = selectedCultures.Select(CultureInfo.GetCultureInfo)
                .ToDictionary(
                    cult => cult.LCID,
                    cult => cult.LCID == InvariantCultureId ? InvariantCultureDisplayName : cult.TwoLetterISOLanguageName.ToUpper()
                );

            var culturesOrder = new List<int>(cultures.Count)
            {
                InvariantCultureId
            };
            culturesOrder.AddRange(cultures.Where(cult => cult.Key != InvariantCultureId).OrderBy(cult => cult.Value).Select(cult => cult.Key));

            var header = new HeaderModel<ResExcelModel>
            {
                Columns = new List<ColumnModel>(1) {new ColumnModel {Title = "Resource Key"}}
                    .Concat(culturesOrder.Select(cultureId => cultures[cultureId]).Select(headerName => new ColumnModel {Title = headerName}))
                    .ToList()
            };

            IReadOnlyList<ResGroupModel<ResExcelModel>> groups = solutionResources
                .ProjectResources.Select(proj => new ResGroupModel<ResExcelModel>
                {
                    GroupTitle = proj.ProjectName,
                    Tables = proj.Resources.Select(res =>
                    {
                        var neutralResources = res.Value[InvariantCultureId].Resources;
                        List<string> keysOrder = neutralResources.Keys.OrderBy(key => key).ToList();

                        List<RowModel<ResExcelModel>> rows = keysOrder.Select(
                            resKey => new RowModel<ResExcelModel>
                            {
                                Model = new ResExcelModel(resKey, culturesOrder.Select(cultureId => res.Value[cultureId]).Select(resData => resData.Resources[resKey]).ToList())
                            })
                            .ToList();

                        var tableModel = new ResTableModel<ResExcelModel>
                        {
                            TableTitle = res.Key,
                            Header = header,
                            Rows = rows
                        };

                        return tableModel;
                    })
                        .ToList()
                })
                .ToList();

            return _excelGenerator.ExportToExcel(groups, title);
        }

        public void ImportFromExcel(IReadOnlyCollection<int> selectedCultures, IReadOnlyCollection<Project> selectedProjects, FileInfoContainer file)
        {
            IReadOnlyList<ResGroupModel<ResExcelModel>> data = _excelGenerator.ImportFromExcel<ResExcelModel>(file);
            SolutionResources resources = GetSolutionResources(selectedCultures, selectedProjects);

            var projectsJoin = resources.ProjectResources
                .Join(data, projRes => projRes.ProjectName, excelProjRes => excelProjRes.GroupTitle, (projRes, excelProjRes) => new {ProjRes = projRes, ExcelProjRes = excelProjRes});

            foreach (var project in projectsJoin)
            {
                var resourceTablesJoin = project.ProjRes.Resources
                    .Join(project.ExcelProjRes.Tables, resTable => resTable.Key, excelResTable => excelResTable.TableTitle, (resTable, excelResTable) => new {ResTable = resTable, ExcelResTable = excelResTable});

                foreach (var resource in resourceTablesJoin)
                {
                    List<int> cultureIds = resource.ExcelResTable.Header.Columns
                        .Skip(1)
                        .Select(col => col.Title == InvariantCultureDisplayName ? InvariantCultureId : CultureInfo.GetCultureInfo(col.Title).LCID)
                        .ToList();
                    List<string> resourceKeys = resource.ExcelResTable.Rows.Select(row => row.DataList[0].DataString).ToList();

                    Dictionary<int, Dictionary<string, string>> excelResources = cultureIds
                        .Select((cultureId, index) => new KeyValuePair<int, Dictionary<string, string>>
                            (
                            cultureId,
                            resourceKeys.Zip(
                                resource.ExcelResTable.Rows.Select(row => row.DataList[index + 1].DataString),
                                (key, value) => new KeyValuePair<string, string>(key, value)
                                )
                                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value)
                            ))
                        .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

                    var resourceFileTablesJoin = resource.ResTable.Value
                        .Join(excelResources, resData => resData.Key, excelResData => excelResData.Key, (resData, excelResData) => new {ResData = resData.Value, ExcelResData = excelResData.Value});

                    foreach (var resFileTablesJoin in resourceFileTablesJoin)
                    {
                        try
                        {
                            UpdateResourceFile(resFileTablesJoin.ResData, resFileTablesJoin.ExcelResData);
                        }
                        catch (MissingManifestResourceException ex)
                        {
                            throw new MissingManifestResourceException(String.Format("There are missing resources in project {0}, file {1} culture {2}",
                                project.ProjRes.ProjectId, resFileTablesJoin.ResData.ResourceName, resFileTablesJoin.ResData.Culture.DisplayName), ex);
                        }
                    }
                }
            }
        }

        private void UpdateResourceFile(ResourceData resData, Dictionary<string, string> excelResData)
        {
            var resourcesJoin = resData.Resources
                .Join(excelResData, res => res.Key, resExcel => resExcel.Key,
                    (res, resExcel) => new { Key = res.Key, Value = res.Value, ExcelValue = resExcel.Value })
                .ToList();
            if (resData.Resources.Count != resourcesJoin.Count)
            {
                throw new MissingManifestResourceException("There are some missed resources");
            }

            if (resourcesJoin.Any(res => res.Value != res.ExcelValue))
            {
                using (var writer = new ResXResourceWriter(resData.ResourcePath))
                {
                    foreach (var keyValuePair in resourcesJoin)
                    {
                        writer.AddResource(keyValuePair.Key, keyValuePair.ExcelValue);
                    }
                }
            }
        }

        private void UpdateResourceFiles(ResourceData neutralCulture, IEnumerable<ResourceData> cultureFiles)
        {
            Dictionary<string, string> neutralCultureResources = neutralCulture.Resources;

            foreach (var resourceFileInfo in cultureFiles)
            {
                Dictionary<string, string> cultResources = resourceFileInfo.Resources;

                if (!cultResources.Keys.OrderBy(k => k).SequenceEqual(neutralCultureResources.Keys.OrderBy(k => k)))
                {
                    using (var writer = new ResXResourceWriter(resourceFileInfo.ResourcePath))
                    {
                        var resources = cultResources.Where(res => neutralCultureResources.ContainsKey(res.Key))
                            .Concat(neutralCultureResources.Where(res => !cultResources.ContainsKey(res.Key)))
                            .ToList();

                        foreach (var keyValuePair in resources)
                        {
                            writer.AddResource(keyValuePair.Key, keyValuePair.Value);
                        }
                    }
                }
            }
        }

        private void UpdateHierarchy(ProjectItem neutralResItem, IReadOnlyCollection<ProjectItem> resItems)
        {
            foreach (var projectItem in resItems.Except(neutralResItem.ProjectItems.Cast<ProjectItem>()).ToList())
            {
                projectItem.Remove();
                neutralResItem.ProjectItems.AddFromFile(projectItem.FileNames[0]);
            }
        }


        private Dictionary<string, string> GetResourceContent(string fileName)
        {
            var cultResources = new Dictionary<string, string>();

            using (var reader = new ResXResourceReader(fileName))
            {
                var enumerator = reader.GetEnumerator();

                while (enumerator.MoveNext())
                {
                    cultResources.Add((string) enumerator.Key, (string) enumerator.Value);
                }
            }

            return cultResources;
        }

        private SolutionResources GetSolutionResources(IEnumerable<int> selectedCultures, IEnumerable<Project> projects)
        {
            Func<ResourceData, bool> resourceFilesFilter;
            if (selectedCultures == null)
            {
                resourceFilesFilter = r => true;
            }
            else
            {
                var selectedCulturesHashSet = new HashSet<int>(selectedCultures);

                resourceFilesFilter = r => selectedCulturesHashSet.Contains(r.Culture.LCID);
            }

            return new SolutionResources
            {
                ProjectResources = projects.Select(project =>
                    new ProjectResources
                    {
                        ProjectName = project.Name,
                        ProjectId = project.UniqueName,

                        Resources = GetProjectItems(project.ProjectItems.Cast<ProjectItem>())
                            .Where(projItem => Path.GetExtension(projItem.FileNames[0]) == ".resx")

                            .Select(projItem =>
                            {
                                string fileName = projItem.FileNames[0];

                                string resName = Path.GetFileNameWithoutExtension(fileName);
                                string cultureName = Path.GetExtension(resName);
                                resName = Path.GetFileNameWithoutExtension(resName);

                                return new ResourceData
                                    (
                                    resourceName: resName,
                                    resourcePath: fileName,
                                    culture: String.IsNullOrEmpty(cultureName) ? CultureInfo.InvariantCulture : CultureInfo.GetCultureInfo(cultureName.Substring(1)),
                                    projectItem: projItem,
                                    resources: GetResourceContent(fileName)
                                    );
                            })
                            .Where(resourceFilesFilter)
                            .GroupBy(res => res.ResourceName)
                            .ToDictionary(resGroup => resGroup.Key, resGroup => resGroup.ToDictionary(res => res.Culture.LCID, res => res))

                    })
                    .ToList()
            };
        }

        private IEnumerable<ProjectItem> GetProjectItems(IEnumerable<ProjectItem> projectItems)
        {
            var projectsList = projectItems as IList<ProjectItem> ?? projectItems.ToList();

            return projectsList.Concat(projectsList.SelectMany(pi => GetProjectItems(pi.ProjectItems.Cast<ProjectItem>())));
        }
    }
}
