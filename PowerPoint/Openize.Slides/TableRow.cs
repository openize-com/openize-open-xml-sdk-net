using Openize.Slides.Common;
using Openize.Slides.Common.Enumerations;
using Openize.Slides.Facade;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Openize.Slides
{
    /// <summary>
    /// Represents a row within a table.
    /// </summary>
    public class TableRow
    {
        private List<TableCell> _Cells;
        private int _ID;
        private int _RowHeight;
        private Stylings _RowStylings;

        /// <summary>
        /// Gets or sets the unique identifier of the row.
        /// </summary>
        public int ID { get => _ID; set => _ID = value; }

        /// <summary>
        /// Gets or sets the height of the row.
        /// </summary>
        public int RowHeight { get => _RowHeight; set => _RowHeight = value; }

        /// <summary>
        /// Gets or sets the list of cells in the row.
        /// </summary>
        public List<TableCell> Cells { get => _Cells; set => _Cells = value; }

        /// <summary>
        /// Gets or sets the stylings applied to the row.
        /// </summary>
        public Stylings RowStylings { get => _RowStylings; set => _RowStylings = value; }

        /// <summary>
        /// Default constructor for the TableRow class. Initializes a new instance of the TableRow class with an empty list of cells.
        /// </summary>
        public TableRow()
        {
            _Cells = new List<TableCell>();
        }

        /// <summary>
        /// Constructor for the TableRow class that initializes a new instance of the TableRow class with a reference to the table's stylings.
        /// </summary>
        /// <param name="table">The table containing the row.</param>
        public TableRow(Table table)
        {
            _RowStylings = table.TableStylings;
            _Cells = new List<TableCell>();
        }

        /// <summary>
        /// Adds a cell to the row.
        /// </summary>
        /// <param name="cell">The TableCell object to be added to the row.</param>
        public void AddCell(TableCell cell)
        {
            // If row stylings are defined, apply them to the cell.
            if (_RowStylings.FontSize > 0)
            {
                cell.CellStylings = _RowStylings;
            }

            // Add the cell to the row.
            _Cells.Add(cell);
        }
    }

}
