using FirebirdSql.Data.FirebirdClient;
using System;
using System.Configuration;
using System.Data;

namespace importerWZ
{
    class DataBaseManger
    {

        private static string connectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ToString();

        public DataTable ExecuteQueryResult(string command)
        {
            DataTable dataTable = new DataTable();
            using (FbConnection connection = new FbConnection(DataBaseManger.connectionString))
            {
                connection.Open();
                FbDataAdapter fbDataAdapter = new FbDataAdapter(new FbCommand(command, connection));
                DataSet dataSet = new DataSet();
                fbDataAdapter.Fill(dataSet);
                dataTable = dataSet.Tables[0];
            }
            return dataTable;
        }

        public void Execute(string command)
        {
            using (FbConnection connection = new FbConnection(DataBaseManger.connectionString))
            {
                connection.Open();
                new FbCommand(command, connection).ExecuteNonQuery();
            }
        }

        public int ExecuteReturnId(string command)
        {
            using (FbConnection connection = new FbConnection(DataBaseManger.connectionString))
            {
                connection.Open();
                FbCommand fbCommand = new FbCommand(command, connection);
                fbCommand.Parameters.Add("Id", FbDbType.Integer).Direction = ParameterDirection.Output;
                return (int)fbCommand.ExecuteScalar();
            }
        }

        public string ExecuteReturnVal(string command)
        {
            using (FbConnection connection = new FbConnection(DataBaseManger.connectionString))
            {
                connection.Open();
                return new FbCommand(command, connection).ExecuteScalar()?.ToString();
            }
        }

        public bool CheckConnection()
        {
            using (FbConnection fbConnection = new FbConnection(DataBaseManger.connectionString))
            {
                try
                {
                    fbConnection.Open();
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            return true;
        }
    }


}