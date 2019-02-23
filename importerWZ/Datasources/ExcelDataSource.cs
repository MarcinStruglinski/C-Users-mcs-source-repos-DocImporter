using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace importerWZ.Datasources
{
    class ExcelDataSource
    {
        private string connectionString;
        private string workSheetName;

        public ExcelDataSource(string path)
        {
            this.connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}; Extended Properties=Excel 12.0;", (object)path);
            this.workSheetName = ConfigurationManager.AppSettings["WorkSheet"].ToString();
        }

        static void PrintDataSet(DataSet ds)
        {
            Console.WriteLine("Tables in '{0}' DataSet.\n", ds.DataSetName);
            foreach (DataTable dt in ds.Tables)
            {
                Console.WriteLine("{0} Table.\n", dt.TableName);
                for (int curCol = 0; curCol < dt.Columns.Count; curCol++)
                {
                    Console.Write(dt.Columns[curCol].ColumnName.Trim() + "\t");
                }
                for (int curRow = 0; curRow < dt.Rows.Count; curRow++)
                {
                    for (int curCol = 0; curCol < dt.Columns.Count; curCol++)
                    {
                        Console.Write(dt.Rows[curRow][curCol].ToString().Trim() + "\t");
                    }
                    Console.WriteLine();
                }
            }
        }



        public DataTable GetSheetAsDataTable()
        {
            using (OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter("SELECT * FROM [" + this.workSheetName + "$]", this.connectionString))
            {
                DataSet dataSet = new DataSet();
                oleDbDataAdapter.Fill(dataSet);
                ///  PrintDataSet(dataSet);
                return dataSet.Tables[0].Rows.Cast<DataRow>().Where<DataRow>((Func<DataRow, bool>)(row => ((IEnumerable<object>)row.ItemArray).Any<object>((Func<object, bool>)(field => !(field is DBNull))))).CopyToDataTable<DataRow>();
            }
        }

    }
}
