using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gs.Excel.Comparer.Structures.Classes
{
    public class ItemRow
    {


        public ItemRow(string phoneNumber)
        {
            PhoneNumber = phoneNumber;
        }

        public string PhoneNumber { get; set; }
    }
}
