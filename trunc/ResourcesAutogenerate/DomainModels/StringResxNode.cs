using System.Reflection;
using System.Resources;

namespace ResourcesAutogenerate.DomainModels
{
    public class StringResxNode
    {
        public StringResxNode(ResXDataNode resXDataNode)
        {
            Name = resXDataNode.Name;
            Value = (string) resXDataNode.GetValue(new AssemblyName[0]);
            Comment = resXDataNode.Comment;
        }

        public string Name { get; set; }
        public string Value { get; set; }
        public string Comment { get; set; }
    }
}
