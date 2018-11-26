using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DocxTemplateParser.Extensions
{
    public static class ObjectExtensions
    {
        public static IEnumerable<string> GetpropertyNames(this object source, BindingFlags bindingAttr = BindingFlags.FlattenHierarchy | BindingFlags.Instance | BindingFlags.Public)
        {
            return source.GetType().GetProperties(bindingAttr).Select(p => p.Name);
        }
        public static IDictionary<string, object> AsDictionary(this object source, BindingFlags bindingAttr = BindingFlags.FlattenHierarchy | BindingFlags.Instance | BindingFlags.Public)
        {
            return source.GetType().GetProperties(bindingAttr).ToDictionary
            (
                propInfo => propInfo.Name,
                propInfo => (propInfo.PropertyType.IsClass && propInfo.PropertyType != typeof(string)) ? propInfo.GetValue(source, null).AsDictionary() : propInfo.GetValue(source, null)

            );

        }

        public static object GetPropValue(this object obj, String propName)
        {
            string[] nameParts = propName.Split('.');
            if (nameParts.Length == 1)
            {
                var prop = obj.GetType().GetProperty(propName);
                if (prop == null) return null;
                return prop.GetValue(obj, null);

            }

            foreach (String part in nameParts)
            {
                if (obj == null) return obj;

                Type type = obj.GetType();
                PropertyInfo info = type.GetProperty(part);
                if (info == null) return null;

                obj = info.GetValue(obj, null);
            }

            return obj;
        }
    }
}
