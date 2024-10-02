using Microsoft.Office.Interop.Excel;
using System.Drawing;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Diagnostics;
using System.Reflection;
using System.Globalization;
namespace Chart_Gantt.Excel
{

    public class ExcelHelper : IDisposable
    {
        private Application _excel;
        private Workbook _workbook;
        private string _filePath;
        private int lastRow;
        private Worksheet _worksheet;
        private Range _range;

        public ExcelHelper()
        {
            _excel = new Application();
        }


        internal bool Open(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    _workbook = _excel.Workbooks.Open(filePath);
                    _filePath = filePath;
                }
                else
                    throw new NotSupportedException();
                _worksheet = _excel.ActiveSheet;
                lastRow = _worksheet.UsedRange.Rows.Count + _worksheet.UsedRange.Row + 1;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        //Заполнить ячейку значением
        internal void SetColumn(int column, int row, object data, bool boldText = false)
        {
            try
            {
                _range = _worksheet.Cells[lastRow + row, column];
                _range.Font.Size = 10;
                if (DateTime.TryParse(data.ToString(), out DateTime valueTime))
                    _range.Value = valueTime;
                else _range.Value = data;

                if(boldText)
                    _range.Font.Bold = true;
            }
            catch (Exception)
            { }
        }

        //Установить формулу  для ячейки "Длительность"
        internal void SetFormula(int column, int row)
        {
            try
            {
                ((Worksheet)_excel.ActiveSheet).Cells[lastRow + row, column].FormulaLocal = $"=F{lastRow + row}-E{lastRow + row}";
            }
            catch (Exception)
            { throw; }
        }

        //Метод заполнения гистограммы
        internal void FillColor(string titleAction, DateTime dateStart, DateTime dateFinish, ExcelHelper helper, int row)
        {
            Color color;
            int columnTime = helper.FindTimeColumn(dateStart);
            DateTime checkTime = dateFinish.Date.Add(new TimeSpan(7, 15, 0));
            while (dateStart <= dateFinish && columnTime > 0 && columnTime <= 47)
            {
                if (titleAction.Contains("ХД. Расчет витрин") || titleAction.Contains("ХД. Инкремент после резервов"))
                    color = dateStart > checkTime.Add(new TimeSpan(2, 0, 0)) && dateStart < checkTime.Add(new TimeSpan(4, 0, 0))
                        ? Color.Yellow : dateStart > checkTime.Add(new TimeSpan(4, 0, 0)) && dateStart < checkTime.Add(new TimeSpan(7, 0, 0))
                        ? Color.Red : Color.FromArgb(0, 176, 80);
                else if (titleAction.Contains("Перенос из ХД в BI"))
                    color = dateStart > checkTime.Add(new TimeSpan(4, 0, 0)) && dateStart < checkTime.Add(new TimeSpan(7, 0, 0))
                        ? Color.Yellow : Color.FromArgb(0, 176, 80);
                else
                    color = dateStart > checkTime.Add(new TimeSpan(-2, 0, 0)) && dateStart < checkTime.Add(new TimeSpan(-1, 0, 0))
                        ? Color.Yellow : dateStart > checkTime.Add(new TimeSpan(-1, 0, 0)) && dateStart < checkTime.Add(new TimeSpan(6, 0, 0))
                        ? Color.Red : Color.FromArgb(0, 176, 80);
                helper.SetColor(columnTime, row, color);
                dateStart = dateStart.AddMinutes(30);
                columnTime++;
            }
        }


        //Заполнить ячейку цветом
        internal void SetColor(int column, int row, Color color)
        {
            try
            {
                ((Worksheet)_excel.ActiveSheet).Cells[lastRow + row, column].Interior.Color = color;
            }
            catch (Exception)
            { throw; }
        }

        //Заполнить границу в отчёте
        internal void SetBorders(int row, DateTime closeDay, DateTime openDay, CultureInfo culture)
        {
            try
            {
                _range = _excel.ActiveSheet.Range($"A{lastRow}:AT{lastRow + row}");
                _range.Borders.Color = Color.Black;
                Range rangeDay = _excel.ActiveSheet.Range($"A{lastRow}:A{lastRow + row}");
                rangeDay.Merge(Type.Missing);
                rangeDay.Value = $"{closeDay.ToString("dd")}-{openDay.ToString("dd")} {openDay.ToString("MMMM", culture)}";
                rangeDay.Interior.Color = Color.FromArgb(196, 215, 155);
            }
            catch (Exception)
            { throw; }
        }
        
        //Установить узор для ячейки
        internal void SetPattern(int column, int row)
        {
            try
            {
                int startCol = 8;
                while (startCol <= column)
                {
                    _range = _worksheet.Cells[lastRow + row, startCol];
                    _range.Interior.Pattern = XlPattern.xlPatternLightDown;
                    startCol++;
                }
                int startRow = lastRow;
                while (startRow <= lastRow + row)
                {
                    _range = _worksheet.Cells[startRow, column];
                    _range.Interior.Pattern = XlPattern.xlPatternLightDown;
                    startRow++;
                }
            }
            catch (Exception)
            { }
        }

        //Найти номер столбца по значению времени
        internal int FindTimeColumn(DateTime dateStart)
        {
            try
            { 
                string str = dateStart.ToString("H:mm:ss");
                Range range = _worksheet.Rows[1].Find(str, Missing.Value, Missing.Value, XlLookAt.xlWhole);
                if (range != null)
                    return range.Column;
                else return 0;
            }
            catch (Exception)
            { return 0;}
        }

        internal void Save() => _workbook.Save();


        void IDisposable.Dispose()
        {
            try
            {
                _workbook.Close();
                _excel.Quit();
                Process[] prs = Process.GetProcessesByName("excel");
                foreach (Process proc in prs)
                    if (proc.MainWindowHandle == IntPtr.Zero) proc.Kill();
                Console.WriteLine("Отчёт Lotus заполен.");
            }
            catch (Exception)
            { }
        }
    }
}
