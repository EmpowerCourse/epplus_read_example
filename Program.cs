using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;

namespace epplus_example
{
    class Program
    {
        private const string FILE_PATH = @"k:\code\c#\epplus_example\docs\read_me.xlsx";

        static void Main(string[] args)
        {
            List<Person> people = readPeople();
            foreach(var p in people)
            {
                Console.WriteLine(String.Format($"Read {p.FirstName} {p.LastName} with email address of {p.Email} from an Excel file"));
            }
            Console.ReadKey();
        }

        private static List<Person> readPeople()
        {
            var result = new List<Person>();
            var fileInfo = new FileInfo(FILE_PATH);
            var excelPackage = new ExcelPackage(fileInfo);
            var excelWorksheet = excelPackage.Workbook.Worksheets["Sheet1"];
            var lastRow = excelWorksheet.Dimension.End.Row;
            // start at row 2 since our data has a header row with titles in it
            for (int rw = 2; rw <= lastRow; rw++)
            {
                var person = new Person()
                {
                    FirstName = parseNullableString(excelWorksheet.Cells[rw, 1].Value),
                    LastName = parseNullableString(excelWorksheet.Cells[rw, 2].Value),
                    Email = parseNullableString(excelWorksheet.Cells[rw, 3].Value)
                };
                result.Add(person);
            }
            return result;
        }

        private static string parseNullableString(object cellValue)
        {
            if (cellValue == null) return null;
            return cellValue.ToString().Trim();
        }
    }
}
