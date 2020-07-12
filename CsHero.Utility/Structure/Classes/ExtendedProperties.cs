using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gs.Utility.Excel.Structure.Classes
{
    [Flags]
    public enum ExtendedProperties
    {
        [StringValue("HDR=No;")]
        HDR_NO = 1,
        [StringValue("HDR=YES;")]
        HDR_YES = 2,
        [StringValue("Excel 12.0 Xml;")]
        Excel_12 = 4,
        [StringValue("Excel 8.0;")]
        Excel_08 = 8,
        [StringValue("IMEX=0;")]
        IMEX_NO = 16,
        [StringValue("IMEX=1;")]
        IMEX_YES = 32
    }
}
