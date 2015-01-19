namespace ResourcesAutogenerate.DomainModels
{
    public class ResourceEntryData
    {
        public ResourceEntryData(string key, string value, string comment)
        {
            Key = key;
            Value = value;
            Comment = comment;
        }

        public string Key { get; set; }

        public string Value { get; set; }

        public string Comment { get; set; }
    }
}
