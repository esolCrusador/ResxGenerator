using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace GloryS.ResxPackage
{
    public class FileChangeTracker : SFileChangeTracker, IVsFileChangeEvents
    {
        public int FilesChanged(uint cChanges, string[] rgpszFile, uint[] rggrfChange)
        {
            MessageBox.Show("FilesChanged");

            return VSConstants.S_OK;
        }

        public int DirectoryChanged(string pszDirectory)
        {
            MessageBox.Show("DirectoryChanged");

            return VSConstants.S_OK;
        }
    }
}
