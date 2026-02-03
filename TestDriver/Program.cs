using System.Diagnostics;
using Microsoft.Office.Interop.Excel;

namespace TestDriver
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Starting TestExcelAddIn");

            PrintExcelInstances("Before starting");
            TestExcelAddIn(test190: false);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Thread.Sleep(5000); // Wait a bit for Excel to exit

            PrintExcelInstances("After TestExcelAddIn with Test180");

            TestExcelAddIn(test190: true);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Thread.Sleep(5000); // Wait a bit for Excel to exit
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Thread.Sleep(5000); // Wait a bit for Excel to exit
            PrintExcelInstances("After TestExcelAddIn with Test190");

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        static void TestExcelAddIn(bool test190)
        {
            // Create a new Excel instance (not shown), and load the add-in from "C:\Temp\TestComRelease\Test180\bin\Debug\net8.0-windows\Test180-AddIn64.xll"
            // using Application.RegisterXLL method.

            var app = new Application();
#if ORIGINAL_CODE
            if (test190)
            {
                app.RegisterXLL(@"C:\Temp\TestComRelease\Test190\bin\Debug\net8.0-windows\Test190-AddIn64.xll");
            }
            else
            {
                app.RegisterXLL(@"C:\Temp\TestComRelease\Test180\bin\Debug\net8.0-windows\Test180-AddIn64.xll");
            }

            // Create a new workbook in the temp directory,
            // on the first sheet, add a formula to cell A1 to call the SayHello function from Test180 add-in.
            // Get the value from that cell and write to debug.
            var workbook = app.Workbooks.Add();
            var sheet = (Worksheet)workbook.Sheets[1];
            sheet.Range["A1"].Formula = "=SayHello(\"Excel\")";
            var result = sheet.Range["A1"].Value;
            Console.WriteLine(result);

            // Close Excel, without saving the workbook
            workbook.Close(SaveChanges: false);
#else
            var Workbooks = app.Workbooks;
            foreach (dynamic rcw in Workbooks)
            {
                rcw.Close(SaveChanges: false);
            }

            var ConfigurationWorkbook = Workbooks.Add();
            app.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;

            // var props = ConfigurationWorkbook.CustomDocumentProperties;

            // var Addins = app.AddIns2;

            // Register the add in
            FileInfo xllInfo;
#if true
            if (test190)
            {
                xllInfo = new FileInfo(@"C:\Temp\TestComRelease\Test190\bin\Debug\net8.0-windows\Test190-AddIn64.xll");
            }
            else
            {
                xllInfo = new FileInfo(@"C:\Temp\TestComRelease\Test180\bin\Debug\net8.0-windows\Test180-AddIn64.xll");
            }
#endif
            if (xllInfo.Exists)
            {
                //var Addin = Addins.Add(xllInfo.FullName);

                //if (Addin == null)
                //    throw new Exception("Unable to find Addin");

                //if (Addin.Installed)
                //{
                //    Addin.Installed = false;
                //    Addin.Installed = true;
                //}
                //else
                //{
                //    Addin.Installed = true;
                //}

                //Addin.Installed = false;
                app.RegisterXLL(xllInfo.FullName);
                ConfigurationWorkbook.Close(SaveChanges: false);
            }
            else
            {
                throw new Exception($"File does not exist {xllInfo}");
            }
#endif
            app.Quit();
        }

        static void PrintExcelInstances(string heading)
        {
            Console.WriteLine(heading);
            var processes = Process.GetProcessesByName("EXCEL");
            if (processes.Length == 0)
            {
                Console.WriteLine("No Excel instances running.");
                return;
            }

            Console.WriteLine("Running Excel instances:");
            Console.WriteLine(new string('-', 40));

            foreach (var p in processes)
            {
                string title = string.IsNullOrWhiteSpace(p.MainWindowTitle) ? "<no title>" : p.MainWindowTitle;
                Console.WriteLine($"PID: {p.Id,-6}  Title: {title}");
            }
            Console.WriteLine(new string('-', 40) + "\n" );
        }
    }
}
