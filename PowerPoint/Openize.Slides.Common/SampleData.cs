using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Openize.Slides.Common
{
    public static class SampleData
    {
        public static DataTable GenerateSampleDataTable ()
        {
            // Create a DataTable with columns
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("ID", typeof(int));
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Email", typeof(string));

            // Add the first row
            DataRow row1 = dataTable.NewRow();
            row1["ID"] = 897;
            row1["Name"] = "Umar";
            row1["Email"] = "umar12@gmail.com";
            dataTable.Rows.Add(row1);

            // Add the second row
            DataRow row2 = dataTable.NewRow();
            row2["ID"] = 897;
            row2["Name"] = "Asif";
            row2["Email"] = "asif@gmail.com";
            dataTable.Rows.Add(row2);

            return dataTable;
        }
    }
}
