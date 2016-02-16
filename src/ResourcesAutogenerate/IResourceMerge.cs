using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Common.Excel;
using Common.Excel.Contracts;
using Common.Excel.Models;
using EnvDTE;
using ResourcesAutogenerate.DomainModels;

namespace ResourcesAutogenerate
{
    public interface IResourceMerge
    {
        void SetLogger(ILogger logger);

        Task UpdateResourcesAsync(IReadOnlyCollection<string> selectedCultures, IReadOnlyCollection<Project> selectedProjects, IStatusProgress progress, CancellationToken cancellationToken, UpdateResourcesOptions options);

        Task ExportToDocumentAsync(IDocumentGenerator documentGenerator, string path, IReadOnlyCollection<string> selectedCultures, IReadOnlyCollection<Project> selectedProjects, IStatusProgress progress, CancellationToken cancellationToken);
        Task ImportFromDocumentAsync(IDocumentGenerator document, string path, IReadOnlyCollection<string> selectedCultures, IReadOnlyCollection<Project> selectedProjects, IStatusProgress progress, CancellationToken cancellationToken);
    }
}
