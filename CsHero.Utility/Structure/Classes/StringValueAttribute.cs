using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gs.Utility.Excel.Structure.Classes
{
    internal class StringValueAttribute : Attribute
    {
        internal StringValueAttribute(string value)
        {
            this.Value = value;
        }

        public string Value { get; }
    }
}
