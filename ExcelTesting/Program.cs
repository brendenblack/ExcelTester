
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTesting
{
    class Program
    {


        static void Main(string[] args)
        {
            Console.WriteLine("Beginning tests...");
            var report = new FileInfo(@"C:\Workspace\MOH\Rate Report 20150927-20160823.xlsm");
            string sheetName = "C2_CAST";
            string password = "salsa";
            DateTime startDate = new DateTime(2016, 7, 3);
            DateTime endDate = new DateTime(2016, 7, 9);

            Console.WriteLine("Parameters:");
            Console.WriteLine($"\tStart date: {startDate.ToShortDateString()}");
            Console.WriteLine($"\tEnd date: {endDate.ToShortDateString()}");
            var tester = new Tester(startDate, endDate);

            Console.WriteLine($"Opening {report.Name}...");

            var start = DateTime.Now;
            using (var package = new ExcelPackage(report, password))
            {
                Console.WriteLine($"Finished in {(DateTime.Now - start).Seconds} seconds");
                start = DateTime.Now;
                Console.WriteLine("Reading sheet in to memory...");
                var sheet = package.Workbook.Worksheets.Where(s => s.Name == sheetName).FirstOrDefault();
                Console.WriteLine($"Finished in {(DateTime.Now - start).Seconds} seconds");

                object[] paramArray = new object[] { sheet };

                var methods = typeof(Tester).GetMethods().Where(m => m.ReturnType == typeof(TestResult)).ToList();
                var totalTests = methods.Count;
                int i = 1;
                foreach (var method in methods)
                {
                    Console.WriteLine($"Running test {i} of {totalTests}...");
                    var result = (TestResult)method.Invoke(tester, paramArray);
                    Console.WriteLine($"Method: {result.Description}");
                    Console.WriteLine($"Found {result.ResultsFound} results in {result.Seconds} seconds");
                    Console.WriteLine("");
                    i++;
                }

            }
            Console.WriteLine("Done. Prese enter to exit.");
            Console.ReadLine();
            // (DateTime.Now - startTime).Seconds



        }
    }
}
