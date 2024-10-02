using System.Data;

namespace Chart_Gantt.Excel
{
    internal class DataTableManager
    {
        private DataTable _dataTable;

        public DataTableManager(DataTable table) => _dataTable = table;

        public void AddRow(ref DataRow newRow, string titleRow, TimeSpan? executionTime, DateTime startTime, DateTime? finishTime)
        {
            newRow = _dataTable.NewRow();
            newRow[1] = titleRow;
            newRow[2] = executionTime;
            newRow[3] = startTime;
            newRow[4] = finishTime;
        }
    }
}
