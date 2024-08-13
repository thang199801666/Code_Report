using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string path = "D:\\Code\\CSharp\\Code_Report\\Code_Report\\bin\\Debug\\Data\\20240603 Simpson and Competitor Code Report Summary Final (LOCKED).xlsx";

            ExcelReader reader = new ExcelReader(path);
            reader.getTableByRange("SST ESRs & ERs");

            Console.ReadKey();
        }
    }
}
