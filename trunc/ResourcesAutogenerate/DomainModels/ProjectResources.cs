using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ResourcesAutogenerate.DomainModels
{
    public class ProjectResources
    {
        public string ProjectId { get; set; }

        public string ProjectName { get; set; }

        public Dictionary<string, Dictionary<int, ResourceData>> Resources { get; set; }
    }
}
