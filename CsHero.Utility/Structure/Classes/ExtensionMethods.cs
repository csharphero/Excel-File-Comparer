using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Gs.Utility.Excel.Structure.Classes
{
    public static class ExtensionMethods
    {
        public static bool IsNumeric(this string strNumber)
        {
            int number = 0;
            return int.TryParse(strNumber, out number);
        }
    }
}
