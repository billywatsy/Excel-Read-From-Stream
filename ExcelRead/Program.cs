using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRead
{
    class Program
    {
        static void Main(string[] args)
        {
            var file = System.IO.File.Open(@"C:\Users\paradigm\Documents\excelManagers.xlsx", System.IO.FileMode.Open);

            var list = ExcelRead.ToDynamicList(file, 6, true);
        }
    }
}
