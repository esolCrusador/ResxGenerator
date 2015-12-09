// Guids.cs
// MUST match guids.h
using System;

namespace GloryS.ResxPackage
{
    static class GuidList
    {
        public const string GuidResxPackagePkgString = "fa894511-2da4-4287-b8ef-f0befadb13ed";
        public const string GuidResxPackageCmdSetString = "3dbf5a40-c0cd-4091-aed3-bac54f4fe6a8";

        public static readonly Guid GuidResxPackageCmdSet = new Guid(GuidResxPackageCmdSetString);

        public const string GuidResxPackageOutputPaneString = "a49e0895-1c13-4986-97e1-6a5a8b4868e7";

        public static readonly Guid GuidResxPackageOutputPane = new Guid(GuidResxPackageOutputPaneString);

    };
}