using System;
using System.Collections.Generic;
using System.Text;

namespace Openize.Slides
{
    /// <summary>
    /// Represents a column within a table.
    /// </summary>
    public class TableColumn
    {
        private string _Name;

        /// <summary>
        /// Gets or sets the name of the column.
        /// </summary>
        public string Name { get => _Name; set => _Name = value; }
    }

}
