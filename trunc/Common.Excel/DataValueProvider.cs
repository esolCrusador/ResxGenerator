using System;
using System.Linq;
using System.Reflection;
using System.Web.UI;

namespace Common.Excel
{
    internal class DataValueProvider
    {
        private readonly PropertyInfo[] _propertyInfos;

        public DataValueProvider(Type dataObjectType)
        {
            _propertyInfos = dataObjectType.GetProperties();
            //dataObjectType.GetNestedTypes(bin)
            //_propertyInfos[0].
        }

        public string GetValue(object dataObject, string propertyName)
        {
            var val = DataBinder.Eval(dataObject, propertyName).ToString();
            //var val = _propertyInfos.First(p => p.Name == propertyName).GetValue(dataObject).ToString();
            return val;
        }
    }
}
