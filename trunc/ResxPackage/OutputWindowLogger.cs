//===================================================================================
// Microsoft patterns & practices
// Composite Application Guidance for Windows Presentation Foundation and Silverlight
//===================================================================================
// Copyright (c) Microsoft Corporation.  All rights reserved.
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY
// OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
// LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
// FITNESS FOR A PARTICULAR PURPOSE.
//===================================================================================
// The example companies, organizations, products, domain names,
// e-mail addresses, logos, people, places, and events depicted
// herein are fictitious.  No association with any real company,
// organization, product, domain name, email address, logo, person,
// places, or events is intended or should be inferred.
//===================================================================================

using System;
using System.Diagnostics;
using System.Globalization;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using ResourcesAutogenerate;

namespace GloryS.ResxPackage
{
    public class OutputWindowLogger : ILogger
    {
        private readonly IVsOutputWindow _outputWindow;

        public OutputWindowLogger(IVsOutputWindow outputWindow)
        {
            this._outputWindow = outputWindow;
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA1806:DoNotIgnoreMethodResults", MessageId = "Microsoft.VisualStudio.Shell.Interop.IVsOutputWindowPane.OutputStringThreadSafe(System.String)")]
        public void Log(string message)
        {
            Guid projectLinkerPaneGuid = GuidList.GuidResxPackageOutputPane;
            IVsOutputWindowPane outPane;
            int hr = _outputWindow.GetPane(ref projectLinkerPaneGuid, out outPane);
            if (ErrorHandler.Succeeded(hr))
            {
                outPane.OutputStringThreadSafe(message + Environment.NewLine);
            }
            else
            {
                Trace.WriteLine(string.Format(CultureInfo.CurrentUICulture, Resources.ResourcesPackageErrorWithPrefix,
                                              message));
            }
        }
    }
}