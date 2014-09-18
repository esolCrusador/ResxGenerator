using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Common.Excel.Models;

namespace ResourcesAutogenerate.DomainModels
{
    public class ResExcelModel : IRowModel
    {
        public ResExcelModel()
        {
            
        }

        public ResExcelModel(string resourceKey, IReadOnlyList<string> resourceValues)
        {
            ResourceKey = resourceKey;
            ResourceValues = resourceValues;

            var dataList = new List<ResCellModel>(ResourceValues.Count + 1);

            dataList.Add(new ResCellModel(ResourceKey));

            dataList
                .AddRange(
                    ResourceValues.Select((v, idx) => new ResCellModel(v)
                    {
                        Hilight = idx != 0 && resourceValues.Count(rv => rv == v) > 1
                    }));

            DataList = dataList;
        }

        public string ResourceKey { get; set; }

        public IReadOnlyList<string> ResourceValues { get; set; }

        public IReadOnlyList<ICellModel> DataList { get; set; }
    }
}
