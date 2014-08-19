using System.Windows;
using GloryS.ResourcesPackage;

namespace GloryS.ResxPackage
{
    public class ResxPackageWindow: Window
    {
        public ResxPackageWindow(ResourcesControl control)
        {
            this.Content = control;
        }
    }
}
