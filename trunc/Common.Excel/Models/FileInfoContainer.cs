using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common.Excel.Models
{
    [Serializable]
    public class FileInfoContainer
    {
        public byte[] Bytes { get; set; }
        public string FileName { get; set; }

        public FileInfoContainer(byte[] bytes, string fileName)
        {
            Bytes = bytes;
            FileName = fileName;
        }
    }
}
