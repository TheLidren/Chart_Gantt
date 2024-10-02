using Oracle.ManagedDataAccess.Client;
using System.Data;

namespace Chart_Gantt
{
    internal class ConnContext : IDisposable
    {
        readonly string _connectionString;
        readonly OracleConnection _oracleConn;
        DataSet _dataSet;

        public ConnContext(string connString)
        {
            _connectionString = connString;
            _oracleConn = new OracleConnection();
            _dataSet = new DataSet();
        }

        public DataTable OpenConntection(string query)
        {
            _oracleConn.ConnectionString = _connectionString;
            _oracleConn.Open();
            OracleCommand command = new(query);
            command.CommandType = CommandType.Text;
            command.Connection = _oracleConn;
            using (OracleDataAdapter dataAdapter = new())
            {
                dataAdapter.SelectCommand = command;
                dataAdapter.Fill(_dataSet);
            }
            return _dataSet.Tables[0];
        }

        public void Dispose()
        {
            _oracleConn.Close();
        }

    }
}
