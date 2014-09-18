using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using System.Windows;
using Common.Excel.Contracts;
using Common.Excel.Implementation;
using EnvDTE;
using GloryS.ResourcesPackage;
using Microsoft.Internal.VisualStudio.PlatformUI;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using ResourcesAutogenerate;
using ResxPackage.Dialog;
using ResxPackage.Resources;

namespace GloryS.ResxPackage
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.GuidResxPackagePkgString)]
    public sealed class ResxPackagePackage : Package
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public ResxPackagePackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }



        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Debug.WriteLine (string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            IVsUIShell uiShell = GetService<IVsUIShell, SVsUIShell>();
            uiShell.EnableModeless(Convert.ToInt32(true));

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(GuidList.GuidResxPackageCmdSet, (int)PkgCmdIdList.ResxPackage);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID );

                mcs.AddCommand( menuItem );
            }

            CreateOutputWindow();
        }

        #endregion

        private TInterface GetService<TInterface, TService>()
            where TInterface : class
            where TService : class
        {
            return this.GetService(typeof(TService)) as TInterface;
        }

        private void CreateOutputWindow()
        {
            const int initiallyVisible = 1;
            const int clearWhenSolutionUnloads = 1;

            var outputPaneGuid = GuidList.GuidResxPackageOutputPane;
            IVsOutputWindow outputWindow = GetService<IVsOutputWindow, SVsOutputWindow>();
            if (outputWindow == null)
                throw new ResxGenException(PackageRes.FailedToCreateOutputWindow);

            IVsOutputWindowPane existingPane;
            if (ErrorHandler.Failed(outputWindow.GetPane(ref outputPaneGuid, out existingPane)) || existingPane == null)
            {
                ErrorHandler.ThrowOnFailure(outputWindow.CreatePane(ref outputPaneGuid,
                                                                    PackageRes.LoggerOutputPaneTitle,
                                                                    initiallyVisible,
                                                                    clearWhenSolutionUnloads)
                    );
            }

        }

        private ResourcesControl CreateDialog()
        {
            IExcelGenerator excelGenerator = new ExcelGenerator();
            IResourceMerge resourceMerge = new ResourcesSchema(excelGenerator);

            DTE dte = (DTE)ServiceProvider.GlobalProvider.GetService(typeof(DTE));
            IVsOutputWindow outputWindow = GetService<IVsOutputWindow, SVsOutputWindow>();

            return new ResourcesControl(resourceMerge, dte.Solution, new OutputWindowLogger(outputWindow), ShowMessageBox);
        }

        private void ShowMessageBox(string title, string text, DialogIcon dialogIcon)
        {
            ShowMessageBox(title, text, dialogIcon, true);
        }

        private void ShowMessageBox(string title, string text, DialogIcon dialogIcon, bool modal)
        {
            IVsUIShell uiShell = GetService<IVsUIShell, SVsUIShell>();
            Guid clsid = Guid.Empty;
            int result;
            int makeModal = (modal ? 1 : 0);

            OLEMSGICON olemIcon;

            switch (dialogIcon)
            {
                case DialogIcon.NoIcon:
                    olemIcon = OLEMSGICON.OLEMSGICON_NOICON;
                    break;
                case DialogIcon.Critical:
                    olemIcon = OLEMSGICON.OLEMSGICON_CRITICAL;
                    break;
                case DialogIcon.Question:
                    olemIcon = OLEMSGICON.OLEMSGICON_QUERY;
                    break;
                case DialogIcon.Warning:
                    olemIcon = OLEMSGICON.OLEMSGICON_WARNING;
                    break;
                case DialogIcon.Info:
                    olemIcon = OLEMSGICON.OLEMSGICON_INFO;
                    break;
                default:
                    throw new ArgumentOutOfRangeException("dialogIcon");
            }

            ErrorHandler.ThrowOnFailure(
                uiShell.ShowMessageBox(0, // Not used but required by api
                                       ref clsid, // Not used but required by api
                                       title,
                                       text,
                                       String.Empty,
                                       0,
                                       OLEMSGBUTTON.OLEMSGBUTTON_OK,
                                       OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                                       olemIcon,
                                       makeModal,
                                       out result
                        )
                    );

        }

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            var solution = GetService<IVsSolution, SVsSolution>();

            // As the "Edit Links" command is added on the Project menu even when no solution is opened, it must
            // be validated that a solution exists (in case it doesn't, GetSolutionInfo() returns 3 nulled strings).
            string s1, s2, s3;
            ErrorHandler.ThrowOnFailure(solution.GetSolutionInfo(out s1, out s2, out s3));
            if (s1 != null && s2 != null && s3 != null)
            {
                IVsUIShell uiShell = GetService<IVsUIShell, SVsUIShell>();
                IntPtr parentHwnd = IntPtr.Zero;

                Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(uiShell.GetDialogOwnerHwnd(out parentHwnd));

                var window = new ResxPackageWindow(CreateDialog()) {WindowStartupLocation = WindowStartupLocation.CenterOwner};

                uiShell.EnableModeless(0);
                try
                {
                    WindowHelper.ShowModal(window, parentHwnd);
                }
                finally
                {
                    //this will take place after the window is closed
                    uiShell.EnableModeless(1);
                }
            }
        }

    }
}
