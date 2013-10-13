using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportDataTableConsole
{
    public class ElectronicsDataTable
    {
        public static DataTable GetData()
        {
            DataTable dt = new DataTable();
            dt.TableName = "electronics";
            dt.Columns.Add("Id", typeof(int));
            dt.Columns.Add("Data", typeof(DateTime));
            dt.Columns.Add("Name", typeof(string));
            dt.Rows.Add(1, DateTime.Now.Date.AddDays(-3), "TV");
            dt.Rows.Add(2, DateTime.Now.Date.AddDays(-1), "PC");
            dt.Rows.Add(3, new DateTime(2003, 11, 5), "Camera");
            return dt;
        }
    }
}
