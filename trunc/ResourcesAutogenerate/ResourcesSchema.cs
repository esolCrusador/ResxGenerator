using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Resources;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;
using Common.Excel;
using Common.Excel.Contracts;
using Common.Excel.Models;
using DocumentFormat.OpenXml.Office2013.PowerPoint.Roaming;
using EnvDTE;
using ResourcesAutogenerate.DomainModels;
using ResxPackage.Resources;

namespace ResourcesAutogenerate
{
    public class ResourcesSchema : IResourceMerge
    {
        private static readonly string InvariantCultureDisplayName = PackageRes.DefaultCulture;
        private static readonly int InvariantCultureId = CultureInfo.InvariantCulture.LCID;
        private static readonly HashSet<string> NonCultureExtensions = new HashSet<string> {".cshtml", ".aspx", ".ascx", ""};

        private ILogger _logger;

        public void SetLogger(ILogger logger)
        {
            _logger = logger;
        }

        public Task UpdateResourcesAsync(IReadOnlyCollection<int> selectedCultures, IReadOnlyCollection<Project> selectedProjects, IStatusProgress progress, CancellationToken cancellationToken, UpdateResourcesOptions options)
        {
            return Task.Run(() => UpdateResources(selectedCultures, selectedProjects, progress, cancellationToken, options), cancellationToken);
        }

        public async Task ExportToDocumentAsync(IDocumentGenerator documentGenerator, string path, IReadOnlyCollection<int> selectedCultures, IReadOnlyCollection<Project> selectedProjects, IStatusProgress progress, CancellationToken cancellationToken)
        {
            progress.Report(StatusRes.GettingProjectsResources);
            SolutionResources solutionResources = await GetSolutionResourcesAsync(selectedCultures, selectedProjects, progress, cancellationToken);

            progress.Report(StatusRes.PreparingResourcesToExport);

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

            var header = new HeaderModel
            {
                Columns = new List<ColumnModel>(1) { new ColumnModel { Title = ExcelRes.ResourceKey } }
                    .Concat(culturesOrder.Select(cultureId => cultures[cultureId]).Select(headerName => new ColumnModel { Title = headerName }))
                    .Concat(new List<ColumnModel>(1) { new ColumnModel { Title = ExcelRes.Comment } })
                    .ToList()
            };

            IReadOnlyList<ResGroupModel<ResExcelModel>> groups = solutionResources
                .ProjectResources.Select(proj => new ResGroupModel<ResExcelModel>
                {
                    GroupTitle = proj.ProjectName,
                    Tables = proj.Resources.Select(res =>
                    {
                        var neutralResources = res.Value[InvariantCultureId].StringResources;
                        List<string> keysOrder = neutralResources.Keys.OrderBy(key => key).ToList();

                        List<RowModel<ResExcelModel>> rows = keysOrder.Select(
                            resKey => new RowModel<ResExcelModel>
                            {
                                Model = new ResExcelModel(resKey,
                                    culturesOrder.Select(cultureId => res.Value[cultureId]).Select(resData => resData.StringResources[resKey].Value).ToList(),
                                    res.Value[InvariantCultureId].StringResources[resKey].Comment)
                            })
                            .Where(r => r.Model.ResourceValues.Count != 0)
                            .ToList();

                        var tableModel = new ResTableModel<ResExcelModel>
                        {
                            TableTitle = res.Key,
                            Header = header,
                            Rows = rows
                        };

                        cancellationToken.ThrowIfCancellationRequested();

                        return tableModel;
                    })
                        .Where(table => table.Rows.Count != 0)
                        .ToList()
                })
                .Where(res => res.Tables.Count != 0)
                .ToList();

            await documentGenerator.ExportToDocumentAsync(path, groups, progress, cancellationToken);
        }

        public Task ImportFromDocumentAsync(IDocumentGenerator documentGenerator, string path, IReadOnlyCollection<int> selectedCultures, IReadOnlyCollection<Project> selectedProjects, IStatusProgress progress, CancellationToken cancellationToken)
        {
            return Task.Run(() => ImportFromDocument(documentGenerator, path, selectedCultures, selectedProjects, progress, cancellationToken), cancellationToken);
        }

        private async Task ImportFromDocument(IDocumentGenerator documentGenerator, string path, IReadOnlyCollection<int> selectedCultures, IReadOnlyCollection<Project> selectedProjects, IStatusProgress progress, CancellationToken cancellationToken)
        {
            IReadOnlyList<ResGroupModel<ResExcelModel>> data = await documentGenerator.ImportFromDocumentAsync<ResExcelModel>(path, progress, cancellationToken);

            progress.Report(StatusRes.GettingProjectsResources);
            SolutionResources resources = await GetSolutionResourcesAsync(selectedCultures, selectedProjects, progress, cancellationToken);

            progress.Report(StatusRes.MergingResources);

            var projectsJoin = resources.ProjectResources
                .Join(data, projRes => projRes.ProjectName, excelProjRes => excelProjRes.GroupTitle, (projRes, excelProjRes) => new { ProjRes = projRes, ExcelProjRes = excelProjRes });

            foreach (var project in projectsJoin)
            {
                var resourceTablesJoin = project.ProjRes.Resources
                    .Join(project.ExcelProjRes.Tables, resTable => resTable.Key, excelResTable => excelResTable.TableTitle, (resTable, excelResTable) => new { ResTable = resTable, ExcelResTable = excelResTable });

                foreach (var resource in resourceTablesJoin)
                {
                    int columnsCount = resource.ExcelResTable.Header.Columns.Count;
                    int culturesCount = columnsCount - 2;

                    List<int> cultureIds = resource.ExcelResTable.Header.Columns
                        .Skip(1)
                        .Select(col => col.Title == InvariantCultureDisplayName ? InvariantCultureId : CultureInfo.GetCultureInfo(col.Title).LCID)
                        .Take(culturesCount)
                        .ToList();
                    List<string> resourceKeys = resource.ExcelResTable.Rows.Select(row => row.DataList[0].DataString).ToList();
                    List<string> comments = resource.ExcelResTable.Rows.Select(row => row.DataList[columnsCount - 1].DataString).ToList();

                    Dictionary<int, IReadOnlyCollection<ResourceEntryData>> excelResources = cultureIds
                        .Select((cultureId, index) => new KeyValuePair<int, IReadOnlyCollection<ResourceEntryData>>
                            (
                            cultureId,
                            resourceKeys.Zip(
                                resource.ExcelResTable.Rows.Select(row => row.DataList[index + 1].DataString),
                                (key, value) => new { Key = key, Value = value })
                                .Zip(comments, (kvp, comment) => new ResourceEntryData(kvp.Key, kvp.Value, comment))
                                .ToList()
                            ))
                        .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

                    var resourceFileTablesJoin = resource.ResTable.Value
                        .Join(excelResources, resData => resData.Key, excelResData => excelResData.Key, (resData, excelResData) => new { ResData = resData.Value, ExcelResData = excelResData.Value });

                    foreach (var resFileTablesJoin in resourceFileTablesJoin)
                    {
                        try
                        {
                            UpdateResourceFile(resFileTablesJoin.ResData, resFileTablesJoin.ExcelResData);
                        }
                        catch (MissingManifestResourceException ex)
                        {
                            throw new MissingManifestResourceException(String.Format(ErrorsRes.MissingResourcesFormat,
                                project.ProjRes.ProjectId, resFileTablesJoin.ResData.ResourceName, resFileTablesJoin.ResData.Culture.DisplayName), ex);
                        }
                    }
                }
            }
        }

        public void UpdateResources(IReadOnlyCollection<int> selectedCultures, IReadOnlyCollection<Project> selectedProjects, IStatusProgress progress, CancellationToken cancellationToken, UpdateResourcesOptions options)
        {
            IReadOnlyDictionary<int, CultureInfo> selectedCultureInfos = selectedCultures.Select(CultureInfo.GetCultureInfo)
                .ToDictionary(cult => cult.LCID, cult => cult);

            IReadOnlyDictionary<string, Project> projectsDictionary = selectedProjects.ToDictionary(proj => proj.UniqueName, proj => proj);

            progress.Report(StatusRes.GettingProjectsResources);
            SolutionResources solutionResources = GetSolutionResources(null, selectedProjects, progress, cancellationToken);

            progress.Report(StatusRes.GeneratingResx);
            cancellationToken.ThrowIfCancellationRequested();

            int resourceFilesProcessed = 0;
            int resourceFilesCount = solutionResources.ProjectResources.Sum(pr => pr.Resources.Count);

            foreach (var projectResources in solutionResources.ProjectResources)
            {
                var project = projectsDictionary[projectResources.ProjectId];

                foreach (Dictionary<int, ResourceData> resourceFileGroup in projectResources.Resources.Values)
                {
                    //Removing culture files without neutral culture file. TODO: find if it is required.
                    ResourceData neutralCulture;
                    if (!resourceFileGroup.TryGetValue(InvariantCultureId, out neutralCulture))
                    {
                        _logger.Log(String.Format(LoggerRes.MissingNeutralCulture, resourceFileGroup.Values.First().ResourcePath, String.Join("\r\n", resourceFileGroup.Values.Select(r => r.ResourcePath))));

                        foreach (var projectItem in resourceFileGroup.Values)
                        {
                            _logger.Log(String.Format(LoggerRes.RemovedFormat, projectItem.ProjectItem.FileNames[0]));

                            projectItem.ProjectItem.Delete();
                        }

                        resourceFileGroup.Clear();
                        project.Save();
                        continue;
                    }

                    var items2Remove = resourceFileGroup.Where(f => !selectedCultureInfos.ContainsKey(f.Key)).ToList();

                    List<KeyValuePair<int, ProjectItem>> projectItems2Remove;
                    if (items2Remove.Count == 0)
                    {
                        projectItems2Remove = new List<KeyValuePair<int, ProjectItem>>(0);
                    }
                    else
                    {
                        projectItems2Remove = items2Remove
                            .Join(projectResources.ResourceProjectItems, f => f.Value.ResourcePath, item => item.FileNames[0], (f, item) => new KeyValuePair<int, ProjectItem>(f.Key, item))
                            .ToList();
                    }

                    var cultures2Add = selectedCultureInfos.Where(cult => !resourceFileGroup.ContainsKey(cult.Key)).Select(cult => cult.Value).ToList();

                    if (options.RemoveNotSelectedCultures)
                    {
                        foreach (var projectItem in projectItems2Remove)
                        {
                            _logger.Log(String.Format(LoggerRes.RemovedFormat, projectItem.Value.FileNames[0]));

                            projectItem.Value.Delete();
                            resourceFileGroup.Remove(projectItem.Key);
                        }
                    }

                    foreach (var cultureInfo in cultures2Add)
                    {
                        string resourcePath = Path.Combine(Path.GetDirectoryName(neutralCulture.ResourcePath), Path.GetFileName(neutralCulture.ResourceName) + "." + cultureInfo.TwoLetterISOLanguageName.ToUpper() + ".resx");

                        using (File.Create(resourcePath)) { }

                        ProjectItem projectItem = project.ProjectItems.AddFromFile(resourcePath);

                        var newFile = new ResourceData
                            (
                            resourceName: neutralCulture.ResourceName,
                            resourcePath: resourcePath,
                            culture: cultureInfo,
                            projectItem: projectItem,
                            resources: new Dictionary<string, ResXDataNode>(0)
                            );

                        resourceFileGroup.Add(cultureInfo.LCID, newFile);

                        _logger.Log(String.Format(LoggerRes.AddedNewResource, newFile.ResourcePath));
                    }

                    List<ResourceData> otherCultureResources = resourceFileGroup.Where(resData => resData.Key != InvariantCultureId).Select(resData => resData.Value).ToList();

                    if (options.EmbeedSubCultures.HasValue)
                    {
                        if (options.EmbeedSubCultures.Value)
                        {
                            EmbeedResources(neutralCulture.ProjectItem, otherCultureResources);
                        }
                    }

                    if (options.UseDefaultCustomTool.HasValue)
                    {
                        Property customToolProperty = neutralCulture.ProjectItem.Properties.Cast<Property>().First(p => p.Name == "CustomTool");

                        customToolProperty.Value = options.UseDefaultCustomTool.Value ? "PublicResXFileCodeGenerator" : "";
                    }

                    if (options.UseDefaultContentType.HasValue)
                    {
                        foreach (var resProjectItem in resourceFileGroup.Values.Select(g=>g.ProjectItem))
                        {
                            Property itemTypeProperty = resProjectItem.Properties.Cast<Property>().First(p => p.Name == "ItemType");

                            itemTypeProperty.Value = options.UseDefaultContentType.Value ? "EmbeddedResource" : "None";
                        }
                    }

                    UpdateResourceFiles(neutralCulture, otherCultureResources);

                    project.Save();

                    progress.Report((int) Math.Round((double) 100*(++resourceFilesProcessed)/resourceFilesCount));
                    cancellationToken.ThrowIfCancellationRequested();
                }
            }
        }

        private void UpdateResourceFile(ResourceData resData, IEnumerable<ResourceEntryData> docResData)
        {
            //We need to compare only string resources.
            var resourcesJoin = ((from res in resData.StringResources
                join tempDocRes in docResData
                    on res.Key equals tempDocRes.Key
                    into resJoin
                from docRes in resJoin.DefaultIfEmpty()
                select new {res.Key, ResourceData = res.Value, DocResourceData = docRes ?? new ResourceEntryData(res.Value.Name, res.Value.Value, res.Value.Comment)}))
                .Select(resJoin => new {resJoin.Key, resJoin.ResourceData.Value, resJoin.ResourceData.Comment, DocValue = resJoin.DocResourceData.Value, DocComment = resJoin.DocResourceData.Comment})
                .ToList();


            if (resData.StringResources.Count != resourcesJoin.Count)
            {
                throw new MissingManifestResourceException(String.Format(ErrorsRes.MissingResourceKeys, String.Join(", ", resData.StringResources.Where(r => resourcesJoin.All(rj => rj.Key != r.Key)).Select(r => "\"" + r.Key + "\""))));
            }

            if (resourcesJoin.Any(res => res.Value != res.DocValue||res.Comment!=res.DocComment))
            {
                _logger.Log(String.Format(LoggerRes.UpdatedContentFormat, resData.ResourcePath));

                using (var writer = new ResXResourceWriter(resData.ResourcePath))
                {
                    //Adding not string resource types.
                    var resourcesData = resData.NotStringResources
                        .Concat(resourcesJoin.Select(resourceWithEntry =>
                            new KeyValuePair<string, ResXDataNode>(resourceWithEntry.Key,
                                new ResXDataNode(resourceWithEntry.Key, resourceWithEntry.DocValue)
                                {
                                    Comment = resourceWithEntry.DocComment
                                })
                            )
                        );

                    foreach (var keyValuePair in resourcesData)
                    {
                        writer.AddResource(keyValuePair.Value);
                    }
                }
            }
        }

        private void UpdateResourceFiles(ResourceData neutralCulture, IEnumerable<ResourceData> cultureFiles)
        {
            IReadOnlyDictionary<string, ResXDataNode> neutralCultureResources = neutralCulture.Resources;

            foreach (var resourceFileInfo in cultureFiles)
            {
                IReadOnlyDictionary<string, ResXDataNode> cultResources = resourceFileInfo.Resources;

                if (neutralCultureResources.Count == 0 || !cultResources.Keys.OrderBy(k => k).SequenceEqual(neutralCultureResources.Keys.OrderBy(k => k)))
                {
                    if (neutralCultureResources.Count != 0)
                    {
                        _logger.Log(String.Format(LoggerRes.UpdatedContentFormat, resourceFileInfo.ResourcePath));
                    }

                    using (var writer = new ResXResourceWriter(resourceFileInfo.ResourcePath))
                    {
                        var resources = cultResources.Where(res => neutralCultureResources.ContainsKey(res.Key))
                            .Concat(neutralCultureResources.Where(res => !cultResources.ContainsKey(res.Key)))
                            .ToList();

                        foreach (var keyValuePair in resources)
                        {
                            writer.AddResource(keyValuePair.Value);
                        }
                    }
                }
            }
        }

        private void EmbeedResources(ProjectItem neutralResItem, IReadOnlyCollection<ResourceData> resItems)
        {
            foreach (var resItem in resItems.Where(res=>res.ProjectItem.Collection.Parent!= neutralResItem).ToList())
            {
                resItem.ProjectItem.Remove();
                resItem.ProjectItem = neutralResItem.ProjectItems.AddFromFile(resItem.ProjectItem.FileNames[0]);
            }
        }

        private Dictionary<string, ResXDataNode> GetResourceContent(string fileName)
        {
            var cultResources = new Dictionary<string, ResXDataNode>();

            using (var reader = new ResXResourceReader(fileName){BasePath = Path.GetDirectoryName(fileName), UseResXDataNodes = true})
            {
                var enumerator = reader.GetEnumerator();

                while (enumerator.MoveNext())
                {
                    cultResources.Add((string) enumerator.Key, (ResXDataNode)enumerator.Value);
                }
            }

            return cultResources;
        }

        private Task<SolutionResources> GetSolutionResourcesAsync(IEnumerable<int> selectedCultures, IReadOnlyCollection<Project> projects, IStatusProgress progress, CancellationToken cancellationToken)
        {
            return Task.Run(() => GetSolutionResources(selectedCultures, projects, progress, cancellationToken), cancellationToken);
        }

        private SolutionResources GetSolutionResources(IEnumerable<int> selectedCultures, IReadOnlyCollection<Project> projects, IStatusProgress progress, CancellationToken cancellationToken)
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

            var progresses = progress.CreateParallelProgresses(0.7, 0.3);
            var dteProjectsProgress = progresses[0];
            var resourceContentProgress = progresses[1];

            double projectsCount = projects.Count;

            var projectResourceItems = new SolutionResources
            {
                ProjectResources = projects.Select((project, index) =>
                {
                    string projectDirectory= Path.GetDirectoryName(project.FullName);
                    var resourceProjectItems = project.GetAllItems().Where(projItem => Path.GetExtension(projItem.FileNames[0]) == ".resx").ToList();

                    dteProjectsProgress.Report(100*(index + 1)/projectsCount);
                    cancellationToken.ThrowIfCancellationRequested();

                    return new ProjectResources
                    {
                        ProjectName = project.Name,
                        ProjectDirectory = projectDirectory,
                        ProjectId = project.UniqueName,
                        ResourceProjectItems = resourceProjectItems
                    };
                })
                    .ToList()
            };

            double resourceFilesCount = projectResourceItems.ProjectResources.Sum(pr => pr.ResourceProjectItems.Count);
            int resourceIndex = 0;

            foreach (var projectResourceItem in projectResourceItems.ProjectResources)
            {
                int projectDirectoryPathLength = projectResourceItem.ProjectDirectory.Length;

                projectResourceItem.Resources = projectResourceItem.ResourceProjectItems
                    .Select(projItem =>
                    {
                        string fileName = projItem.FileNames[0];

                        //Removing .resx extension.
                        string resName = Path.GetFileNameWithoutExtension(fileName);

                        string cultureName = Path.GetExtension(resName);
                        if (NonCultureExtensions.Contains(cultureName))
                        {
                            cultureName = string.Empty;
                        }
                        else
                        {
                            //Removing culture extension.
                            resName = Path.GetFileNameWithoutExtension(resName);
                        }

                        string directoryName = (Path.GetDirectoryName(fileName) ?? string.Empty).Substring(projectDirectoryPathLength);
                        //Relative path to the resource.
                        resName = Path.Combine(directoryName, resName);

                        resourceContentProgress.Report(100*(++resourceIndex)/resourceFilesCount);
                        cancellationToken.ThrowIfCancellationRequested();

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
                    .ToDictionary(resGroup => resGroup.Key, resGroup => resGroup.ToDictionary(res => res.Culture.LCID, res => res));
            }



            return projectResourceItems;
        }
    }
}
