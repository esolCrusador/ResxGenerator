using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using GloryS.ResourcesPackage;
using Microsoft.VisualStudio.Shell;
using ResxPackage.Resources;

namespace GloryS.ResxPackage
{
    public class ResxPackageWindow : Window
    {
        private readonly ResourcesControl _control;

        public ResxPackageWindow(ResourcesControl control)
        {
            _control = control;
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
