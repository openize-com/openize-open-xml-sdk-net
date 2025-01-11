using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Openize.Cells
{
    public class Range
    {
        private readonly Worksheet _worksheet;

        public uint StartRowIndex { get; }
        public uint StartColumnIndex { get; }
        public uint EndRowIndex { get; }
        public uint EndColumnIndex { get; }

        /// <summary>
        /// Gets the count of columns in the range.
        /// This property calculates the column count based on the start and end column indices.
        /// </summary>
        /// <value>
        /// The count of columns in the range.
        /// </value>
        /// <exception cref="InvalidOperationException">
        /// Thrown if the end column index is less than the start column index.
        /// </exception>
        public uint ColumnCount
        {
            get 
            {
                if (EndColumnIndex < StartColumnIndex)
                {
                    throw new InvalidOperationException("End column index cannot be less than start column index.");
                }
                return EndColumnIndex - StartColumnIndex + 1; 
            }
        }

        /// <summary>
        /// Gets the count of rows in the range.
        /// This property calculates the row count based on the start and end row indices.
        /// </summary>
        /// <value>
        /// The count of rows in the range.
        /// </value>
        /// <exception cref="InvalidOperationException">
        /// Thrown if the end row index is less than the start row index.
        /// </exception>
        public uint RowCount
        {
            get
            {
                if (EndRowIndex < StartRowIndex)
                {
                    throw new InvalidOperationException("End row index cannot be less than start row index.");
                }
                return EndRowIndex - StartRowIndex + 1;
            }
        }

        /// <summary>
        /// Initializes a new instance of the Range class.
        /// </summary>
        /// <param name="worksheet">The worksheet to which this range belongs.</param>
        /// <param name="startRowIndex">The starting row index of the range.</param>
        /// <param name="startColumnIndex">The starting column index of the range.</param>
        /// <param name="endRowIndex">The ending row index of the range.</param>
        /// <param name="endColumnIndex">The ending column index of the range.</param>
        /// <exception cref="ArgumentNullException">
        /// Thrown if the provided worksheet is null.
        /// </exception>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown if the start or end row/column indices are out of range.
        /// </exception>
        public Range(Worksheet worksheet, uint startRowIndex, uint startColumnIndex, uint endRowIndex, uint endColumnIndex)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            StartRowIndex = startRowIndex;
            StartColumnIndex = startColumnIndex;
            EndRowIndex = endRowIndex;
            EndColumnIndex = endColumnIndex;
        }

        /// <summary>
        /// Sets the specified value to all cells within the range.
        /// </summary>
        /// <param name="value">The value to set in each cell of the range.</param>
        /// <exception cref="InvalidOperationException">
        /// Thrown if any cell within the range is not properly initialized or accessible.
        /// </exception>
        public void SetValue(string value)
        {
            for (uint row = StartRowIndex; row <= EndRowIndex; row++)
            {
                for (uint column = StartColumnIndex; column <= EndColumnIndex; column++)
                {
                    var cellReference = $"{ColumnIndexToLetter(column)}{row}";
                    var cell = _worksheet.GetCell(cellReference);
                    cell.PutValue(value);
                }
            }
        }

        /// <summary>
        /// Clears the values and resets the style of all cells within the range.
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// Thrown if any cell within the range is not properly initialized or accessible.
        /// </exception>
        public void ClearCells()
        {
            for (uint row = StartRowIndex; row <= EndRowIndex; row++)
            {
                for (uint column = StartColumnIndex; column <= EndColumnIndex; column++)
                {
                    var cellReference = $"{ColumnIndexToLetter(column)}{row}";
                    var cell = _worksheet.GetCell(cellReference);
                    if (cell != null)
                    {
                        cell.PutValue(string.Empty); // Clearing the value
                        cell.ApplyStyle(0); // Resetting the style if needed
                    }
                }
            }
        }

        /// <summary>
        /// Merges all cells within the range into a single cell.
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// Thrown if the merge operation fails, for example, if the cells cannot be merged.
        /// </exception>
        public void MergeCells()
        {
            string startCellReference = $"{ColumnIndexToLetter(StartColumnIndex)}{StartRowIndex}";
            string endCellReference = $"{ColumnIndexToLetter(EndColumnIndex)}{EndRowIndex}";

            _worksheet.MergeCells(startCellReference, endCellReference);
        }

        /// <summary>
        /// Adds dropdown list validation with the specified options to all cells within the range.
        /// </summary>
        /// <param name="options">The list of options for the dropdown validation.</param>
        /// <exception cref="ArgumentNullException">
        /// Thrown if the options array is null or empty.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown if adding dropdown validation fails for any cell within the range.
        /// </exception>
        public void AddDropdownListValidation(string[] options)
        {
            

            for (uint row = StartRowIndex; row <= EndRowIndex; row++)
            {
                for (uint column = StartColumnIndex; column <= EndColumnIndex; column++)
                {
                    var cellReference = $"{ColumnIndexToLetter(column)}{row}";
                    _worksheet.AddDropdownListValidation(cellReference, options);
                }
            }
        }

        public void ApplyValidation(ValidationRule rule)
        {
            // Iterate over all cells in the range and apply the validation rule
            for (uint row = StartRowIndex; row <= EndRowIndex; row++)
            {
                for (uint column = StartColumnIndex; column <= EndColumnIndex; column++)
                {
                    var cellReference = $"{ColumnIndexToLetter(column)}{row}";
                    _worksheet.ApplyValidation(cellReference, rule);
                }
            }
        }



        private static string ColumnIndexToLetter(uint columnIndex)
        {
            string columnLetter = string.Empty;
            while (columnIndex > 0)
            {
                columnIndex--;
                columnLetter = (char)('A' + columnIndex % 26) + columnLetter;
                columnIndex /= 26;
            }
            return columnLetter;
        }
    }


}
