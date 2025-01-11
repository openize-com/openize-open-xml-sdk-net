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
    /// Represents a cell within a table row.
    /// </summary>
    public class TableCell
    {
        private string _Text;
        private string _FontFamily;
        private int _FontSize;
        private string _ID;
        private Stylings _CellStylings;

        /// <summary>
        /// Gets or sets the text content of the cell.
        /// </summary>
        public string Text { get => _Text; set => _Text = value; }

        /// <summary>
        /// Gets or sets the font family of the text in the cell.
        /// </summary>
        public string FontFamily { get => _FontFamily; set => _FontFamily = value; }

        /// <summary>
        /// Gets or sets the font size of the text in the cell.
        /// </summary>
        public int FontSize { get => _FontSize; set => _FontSize = value; }

        /// <summary>
        /// Gets or sets the unique identifier of the cell.
        /// </summary>
        public string ID { get => _ID; set => _ID = value; }

        /// <summary>
        /// Gets or sets the stylings applied to the cell.
        /// </summary>
        public Stylings CellStylings { get => _CellStylings; set => _CellStylings = value; }

        /// <summary>
        /// Default constructor for the TableCell class.
        /// </summary>
        public TableCell()
        {

        }

        /// <summary>
        /// Constructor for the TableCell class that initializes a new instance of the TableCell class with a reference to the row's stylings.
        /// </summary>
        /// <param name="row">The table row containing the cell.</param>
        public TableCell(TableRow row)
        {
            _CellStylings = row.RowStylings;
        }
    }

}
