using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using GloryS.ResourcesPackage;
using ResxPackage.Resources;

namespace GloryS.ResxPackage
{
    public class ResxPackageWindow: Window
    {
        public ResxPackageWindow(ResourcesControl control)
        {
            this.Width = 800;
            this.Height = 600;

            this.MinWidth = 400;
            this.MinHeight = 300;

            this.Content = control;

            this.Icon = PackageRes.ResxGenIcon.GetImageSource();
            this.Title = PackageRes.Title;
        }
    }
}
