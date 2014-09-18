using System.Globalization;

namespace ResourcesAutogenerate
{
    public class ResourceFileInfo
    {
        public string ResName { get; set; }

        public CultureInfo Culture { get; set; }

        public string FileName { get; set; }
    }
}
