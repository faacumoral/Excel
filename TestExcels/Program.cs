using ExcelReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace TestExcels
{
    class Program
    {
        static void Main(string[] args)
        {
            var file = File.ReadAllBytes("archivo.xlsx");
            ExcelReader.ExcelReader excel = new ExcelReader.ExcelReader(file);

            Console.ReadKey();

        }
    }
}
