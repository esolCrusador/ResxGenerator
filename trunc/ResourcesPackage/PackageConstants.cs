// Guids.cs
// MUST match guids.h
using System;

namespace GloryS.ResourcesPackage
{
    static class PackageConstants
    {
        public const string guidResourcesPackagePkgString = "8694fc53-b09e-4a00-876a-912464ed12d0";
        public const string guidResourcesPackageCmdSetString = "934772ed-bcc1-492e-a255-658019d7b70d";
        public const string guidToolWindowPersistanceString = "9dd4acb2-a9e3-4dd9-8e60-e59cdae1f617";

        public const string guidResourcesPackageOutputPaneString = "a49e0895-1c13-4986-97e1-6a5a8b4868e7";

        public static readonly Guid guidResourcesPackageCmdSet = new Guid(guidResourcesPackageCmdSetString);
        public static readonly Guid guidResourcesPackageOutputPane = new Guid(guidResourcesPackageOutputPaneString);
    };
}