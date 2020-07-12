using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gs.Excel.Comparer.Structures.Classes
{
    public class ExcelFileItem
    {
        public string FilePath { get; set; }
        public string SheetName { get; set; }
        public int PhoneNumberIndex { get; set; }
        public bool IsChecked { get; set; }

    }
}
