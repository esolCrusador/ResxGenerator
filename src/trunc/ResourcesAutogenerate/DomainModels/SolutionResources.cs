using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ResourcesAutogenerate.DomainModels
{
    public class SolutionResources
    {
        public IReadOnlyList<ProjectResources> ProjectResources { get; set; }
    }
}
