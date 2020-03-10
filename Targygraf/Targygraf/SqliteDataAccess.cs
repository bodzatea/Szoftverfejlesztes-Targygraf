using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Targygraf
{
    public class SqliteDataAccess
    {
        private SQLiteConnection sql_conn;
        private SQLiteCommand sql_cmd;
        private SQLiteDataAdapter DB;
        private DataSet DS = new DataSet();
        private DataTable DT = new DataTable();

        //set connection
        public void SetConnection() {
           // sql_conn = new SQLiteConnection("Data Source=TargyakDB.db;Version=3;New=False;Compress=True");
            sql_conn = new SQLiteConnection(LoadConnectionString());

        }

        //set executequery
        public void ExecuteQuery(string txtQuery) {
            SetConnection();
            sql_conn.Open();
            sql_cmd = sql_conn.CreateCommand();
            sql_cmd.CommandText = txtQuery;
            sql_cmd.ExecuteNonQuery();
            sql_conn.Close();
        }


        private static string LoadConnectionString(string id = "Default")
        {
            return ConfigurationManager.ConnectionStrings[id].ConnectionString;
        }
    }

  
}
