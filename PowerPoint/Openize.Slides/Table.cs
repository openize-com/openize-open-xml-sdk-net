using DocumentFormat.OpenXml.Drawing;
using Openize.Slides.Common;
using Openize.Slides.Facade;
using System;
using System.Collections.Generic;
using System.Data;

namespace Openize.Slides
{
    /// <summary>
    /// This class is responsible to create table in a PPT/PPTX presentataion.
    /// </summary>
    public class Table
    {

        private TableFacade _Facade;
        private int _TableIndex = 0;
        private String _Name;
        private double _x;
        private double _y;
        private double _Width;
        private double _Height;
        private List<TableRow> _Rows;
        private List<TableColumn> _Columns;
        private Stylings _TableStylings;
        private String _Theme = null;



        /// <summary>
        /// Property to get or set the TableFacade instance.
        /// </summary>
        public TableFacade Facade { get => _Facade; set => _Facade = value; }
        /// <summary>
        /// Property to get or set the index of a table within slide.
        /// </summary>
        public int TableIndex { get => _TableIndex; set => _Facade.TableIndex = _TableIndex = value; }
        /// <summary>
        /// Property to get or set the table name within the slide.
        /// </summary>
        public string Name { get => _Name; set => _Name = value; }
        /// <summary>
        /// Property to get or set the X coordinate of a table.
        /// </summary>
        public double X { get => _x; set => _x = value; }
        /// <summary>
        /// Property to get or set the Y coordinate of a table.
        /// </summary>
        public double Y { get => _y; set => _y = value; }
        /// <summary>
        /// Property to get or set the width of a table.
        /// </summary>
        public double Width { get => _Width; set => _Width = value; }
        /// <summary>
        /// Property to get or set the height of a table.
        /// </summary>
        public double Height { get => _Height; set => _Height = value; }

        /// <summary>
        /// Property to get or set the list of rows in the table.
        /// </summary>
        public List<TableRow> Rows { get => _Rows; set => _Rows = value; }

        /// <summary>
        /// Property to get or set the list of columns in the table.
        /// </summary>
        public List<TableColumn> Columns { get => _Columns; set => _Columns = value; }

        /// <summary>
        /// Property to get or set the stylings for the table.
        /// </summary>
        public Stylings TableStylings { get => _TableStylings; set => _TableStylings = value; }
        
        /// <summary>
        /// Property to define theme of a table. It can be check in PowerPoint table designs.  
        /// </summary>
        public string Theme { get => _Theme; set => _Theme = value; }

        /// <summary>
        /// Constructor for the Table class. Initializes a new instance of the Table class with empty lists for rows and columns.
        /// </summary>

        public Table()
        {
            _Rows = new List<TableRow>();
            _Columns = new List<TableColumn>();
        }

        /// <summary>
        /// Adds a row to the table.
        /// </summary>
        /// <param name="row">The TableRow object to be added to the table.</param>
        public void AddRow(TableRow row)
        {
            // If table stylings are defined, apply them to the row.
            if (_TableStylings.FontSize > 0)
            {
                row.RowStylings = _TableStylings;
            }

            // Add the row to the table.
            _Rows.Add(row);
        }
        /// <summary>
        /// Adds a row to the table.
        /// </summary>
        /// <param name="row">The TableRow object to be added to the table.</param>
        public void AddColumn(TableColumn column) { }
        /// <summary>
        /// Method to update an existing table.
        /// </summary>
        public void Update()
        {
            _Facade.TableStyle = _Theme;
            _Facade.UpdateTable(GetDataTable());

        }
        /// <summary>
        /// Method to get datatable from openize table to send to facade.
        /// </summary>        
        /// <returns></returns>
        public DataTable GetDataTable()
        {
            DataTable dtable = new DataTable();

            // Adding columns based on TableColumn information
            foreach (TableColumn column in _Columns)
            {
                dtable.Columns.Add(column.Name, typeof(string));
            }

            // Adding rows based on TableRow and TableCell information
            foreach (TableRow row in _Rows)
            {
                DataRow dataRow = dtable.NewRow();

                // Assuming each TableCell in the TableRow corresponds to a column in the DataTable
                foreach (TableCell cell in row.Cells)
                {
                    // Find the corresponding column by matching the cell's position
                    // Assuming the order of columns in _Columns corresponds to the order of cells in _Cells
                    int columnIndex = _Columns.FindIndex(col => col.Name == cell.ID);

                    if (columnIndex >= 0)
                    {
                        // Add cell value to the corresponding column in the DataRow
                        dataRow[columnIndex] = cell.Text;
                        string stylingInfo = Utility.SerializeStyling(cell.CellStylings);
                        dataRow[columnIndex] += ";" + stylingInfo;
                    }
                    else
                    {
                        // Handle the case where the column for the cell is not found
                        // You may want to log a warning or handle it based on your requirements
                        Console.WriteLine($"Column for cell FontFamily {cell.FontFamily} not found in the table.");
                    }
                }

                // Add the populated DataRow to the DataTable
                dtable.Rows.Add(dataRow);
            }

            return dtable;
        }
        /// <summary>
        /// This method is responsible to get the list of Tables
        /// </summary>
        /// <param name="tableFacades"></param>
        /// <returns></returns>
        public static List<Table> GetTables(List<TableFacade> tableFacades)
        {
            List<Table> tables = new List<Table>();
            foreach (var facade in tableFacades)
            {
                Table table = new Table();
                table.Facade = facade;

                // Iterate through columns and add them to the table
                foreach (System.Data.DataColumn column in facade.SD_DTable.Columns)
                {
                    TableColumn tableColumn = new TableColumn();
                    tableColumn.Name = column.ColumnName;
                    table.Columns.Add(tableColumn);
                }

                // Iterate through rows and add them to the table
                foreach (System.Data.DataRow row in facade.SD_DTable.Rows)
                {
                    TableRow tableRow = new TableRow();
                    int columnIndex = 0; // Keep track of column index
                    foreach (var item in row.ItemArray)
                    {
                        TableCell cell = new TableCell();
                        cell.ID = table.Columns[columnIndex].Name; // Set cell ID with column name
                        cell.Text = item.ToString();
                        tableRow.Cells.Add(cell);
                        columnIndex++; // Move to the next column
                    }
                    table.Rows.Add(tableRow);
                }

                tables.Add(table);
            }

            return tables;
        }
        /// <summary>
        /// Inner class representing different table styles.
        /// </summary>
        public static class TableStyle
        {
            public static string LightStyle1 { get; } = "LightStyle1";
            public static string LightStyle2 { get; } = "LightStyle2";
            public static string LightStyle3 { get; } = "LightStyle3";
            public static string LightStyle4 { get; } = "LightStyle4";
            public static string LightStyle5 { get; } = "LightStyle5";
            public static string LightStyle6 { get; } = "LightStyle6";
            public static string LightStyle7 { get; } = "LightStyle7";
            public static string LightStyle8 { get; } = "LightStyle8";
            public static string LightStyle9 { get; } = "LightStyle9";
            public static string LightStyle10 { get; } = "LightStyle10";
            public static string LightStyle11 { get; } = "LightStyle11";
            public static string LightStyle12 { get; } = "LightStyle12";
            public static string LightStyle13 { get; } = "LightStyle13";
            public static string LightStyle14 { get; } = "LightStyle14";
            public static string MediumStyle1 { get; } = "MediumStyle1";
            public static string MediumStyle2 { get; } = "MediumStyle2";
            public static string MediumStyle3 { get; } = "MediumStyle3";
            public static string MediumStyle4 { get; } = "MediumStyle4";
            public static string MediumStyle5 { get; } = "MediumStyle5";
            public static string MediumStyle6 { get; } = "MediumStyle6";
            public static string MediumStyle7 { get; } = "MediumStyle7";
            public static string MediumStyle8 { get; } = "MediumStyle8";
            public static string MediumStyle9 { get; } = "MediumStyle9";
            public static string MediumStyle10 { get; } = "MediumStyle10";
            public static string MediumStyle11 { get; } = "MediumStyle11";
            public static string MediumStyle12 { get; } = "MediumStyle12";
            public static string DarkStyle1 { get; } = "DarkStyle1";
            public static string DarkStyle2 { get; } = "DarkStyle2";
            public static string DarkStyle3 { get; } = "DarkStyle3";
            public static string DarkStyle4 { get; } = "DarkStyle4";
            public static string DarkStyle5 { get; } = "DarkStyle5";
            public static string DarkStyle6 { get; } = "DarkStyle6";
            public static string DarkStyle7 { get; } = "DarkStyle7";
            public static string DarkStyle8 { get; } = "DarkStyle8";
            public static string DarkStyle9 { get; } = "DarkStyle9";
            public static string DarkStyle10 { get; } = "DarkStyle10";
            public static string DarkStyle11 { get; } = "DarkStyle11";
            public static string DarkStyle12 { get; } = "DarkStyle12";
        }


    }
}
