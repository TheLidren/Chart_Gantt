using Chart_Gantt.Excel;
using Chart_Gantt.Extentions;
using System.Data;
using System.Globalization;
namespace Chart_Gantt
{
    internal class Program
    {
        static string _connectionString = "connection string Oracle";
        static string _path = "~\\Исходник\\Отчёт_Gant.xlsx";
        static readonly string _query = File.ReadAllText(Path.Combine(Directory.GetCurrentDirectory(), @"ОтчётLotuscvb.sql"));
        static DataTable _dataTable = new();
        static DataRow _newRowTable = null!;
        static string _titleAction = null!;
        static short count = 0;
        static short countClose = 1;
        static CultureInfo cultureRu = CultureInfo.CreateSpecificCulture("ru-RU");

        private static DataRow SetCloseRow(DataTable dt, DataRow openRow)
        {
            DataRow closeRow = dt.AsEnumerable().Where(r => r.Field<string>("ACTION") == "Закрытие дня. По 117").First();
            DataRow newRow = dt.NewRow();
            newRow[0] = closeRow.ItemArray[0].ToString();
            newRow[1] = closeRow.ItemArray[1].ToString();
            newRow[2] = closeRow.ItemArray[2].ToString();
            newRow[3] = closeRow.ItemArray[3].ToString();
            newRow[4] = closeRow.ItemArray[4].ToString();
            dt.Rows.InsertAt(newRow, dt.Rows.IndexOf(openRow));
            dt.Rows.Remove(closeRow);
            return newRow;
        }

        static void Main(string[] args)
        {
            try
            {
                using (ConnContext conn = new(_connectionString))
                {
                    _dataTable = conn.OpenConntection(_query);
                }
                DataTableManager tableManager = new(_dataTable);
                DataRow openRow = _dataTable.AsEnumerable().Where(r => r.Field<string>("ACTION") == "Открытие дня. Регламент 1").FirstOrDefault();
                DataRow closeRow = SetCloseRow(_dataTable, openRow);
                DateTime closeDay = DateTime.Parse(closeRow.ItemArray[0].ToString(), CultureInfo.CurrentCulture);
                DateTime openDay = DateTime.Parse(openRow.ItemArray[0].ToString(), CultureInfo.CurrentCulture);
                using (var helper = new ExcelHelper())
                {
                    if (helper.Open(filePath: _path))
                    {
                        for (short i = 0; i < _dataTable.Rows.Count; i++)
                        {
                            for (short j = 1; j < _dataTable.Columns.Count; j++)
                            {
                                var value = _dataTable.Rows[i].ItemArray[j].ToString();
                                if (j == 1)
                                {
                                    _titleAction = value;
                                    helper.SetColumn(2, i, $"{++count}");
                                    helper.SetColumn(j + 2, i, value);
                                }
                                else if (j == 2)
                                    helper.SetFormula(j + 2, i);
                                else if (j == 3)
                                    if (DateTime.TryParse(value, out DateTime dateStart))
                                    {
                                        helper.SetColumn(j + 2, i, dateStart);
                                        //Заполнение гистограммы
                                        DateTime dateFinish = string.IsNullOrEmpty(_dataTable.Rows[i].ItemArray[j + 1].ToString()) ? openDay.Date.AddHours(10) : DateTime.Parse(_dataTable.Rows[i].ItemArray[j + 1].ToString());
                                        helper.FillColor(_titleAction, dateStart.AroundMethod(), dateFinish.AroundMethod(), helper, i);
                                    }
                                else
                                    helper.SetColumn(j + 2, i, value);
                            }
                        }
                        //Выводить месяца в Р.п.
                        cultureRu.DateTimeFormat.MonthNames = cultureRu.DateTimeFormat.MonthGenitiveNames;
                        helper.SetBorders(_dataTable.Rows.Count - 1, closeDay, openDay, cultureRu);
                        helper.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "\n Отчёт по диаграмме Ганта не заполнен. Необходимо заполнить вручную!");
            }
            Console.Write("Нажмите Enter для выхода:");
            Console.ReadLine();
        }

    }
}
