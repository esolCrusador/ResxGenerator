namespace ResourcesAutogenerate.DomainModels
{
    public class UpdateResourcesOptions
    {
        public bool RemoveNotSelectedCultures { get; set; }

        public bool? EmbeedSubCultures { get; set; }

        public bool? UseDefaultContentType { get; set; }

        public bool? UseDefaultCustomTool { get; set; }
    }
}
