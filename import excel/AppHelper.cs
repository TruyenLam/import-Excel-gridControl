using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace import_excel
{
    public class AppHelper
    {
        private OleDbConnection cn = null;
        public AppHelper(string connectionString) { cn = new OleDbConnection(connectionString); }

        private void OpenConnection()
        {
            if (cn.State == ConnectionState.Closed)
                cn.Open();
        }

        private void CloseConnection()
        {
            if (cn.State == ConnectionState.Open)
                cn.Close();
        }

        public DataTable GetOleDbSchemaTable()
        {
            DataTable dt = new DataTable();
            try
            {
                OpenConnection();
                dt = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            }
            catch { return null; }
            finally
            {
                CloseConnection();
            }
            return dt;
        }
    }
}
