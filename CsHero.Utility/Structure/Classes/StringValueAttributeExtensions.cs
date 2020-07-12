using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Gs.Utility.Excel.Structure.Classes
{
    internal static class StringValueAttributeExtensions
    {
        internal static string GetStringValue(this Enum value)
        {
            // Get the type
            Type type = value.GetType();
            IEnumerable<ExtendedProperties> values = Enum.GetValues(type).Cast<ExtendedProperties>();

            string result = string.Empty;

            foreach (var item in values)
            {
                if (!value.HasFlag(item))
                    continue;

                // Get fieldinfo for this type
                FieldInfo fieldInfo = type.GetField(item.ToString());

                // Get the stringvalue attributes
                StringValueAttribute[] attribs = fieldInfo.GetCustomAttributes(
                    typeof(StringValueAttribute), false) as StringValueAttribute[];

                result += attribs[0].Value;
            }
            if (string.IsNullOrWhiteSpace(result))
                return string.Empty;
            return result.Substring(0, result.Length - 1);
        }
    }
}
