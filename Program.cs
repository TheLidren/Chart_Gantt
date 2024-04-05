using Microsoft.Office.Interop.Excel;
using System.Reflection;
namespace Report_Lotus
{
    internal class Program
    {
        private static Application oXL;
        private static _Workbook oWB;
        private static _Worksheet oSheet;
        private static Microsoft.Office.Interop.Excel.Range oRng;
        static void Main(string[] args)
        {
            oXL = new Application();
            oXL.Visible = true;

            oWB = (_Workbook)(oXL.Workbooks.Add(Missing.Value));
            oSheet = (_Worksheet)oWB.ActiveSheet;

            oSheet.Cells[1, 1] = "First Name";
            oSheet.Cells[1, 2] = "Last Name";
            oSheet.Cells[1, 3] = "Full Name";
            oSheet.Cells[1, 4] = "Salary";
            Console.WriteLine("Hello, World!");
        }
    }
}
