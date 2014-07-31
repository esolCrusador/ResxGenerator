using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace Common.Excel.Models
{
    public interface ICellModel
    {
        bool Hilight { get; set; }

        string DataString { get; set; }
    }
}
