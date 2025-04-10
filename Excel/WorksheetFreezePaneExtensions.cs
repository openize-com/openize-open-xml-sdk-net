using System;

namespace Openize.Cells
{
    /// <summary>
    /// Provides extension methods for easy management of freeze panes in Excel worksheets.
    /// </summary>
    public static class WorksheetFreezePaneExtensions
    {
        /// <summary>
        /// Freezes the top row of the worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <returns>The worksheet for method chaining.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        public static Worksheet FreezeTopRow(this Worksheet worksheet)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            worksheet.FreezePane(1, 0);
            return worksheet;
        }

        /// <summary>
        /// Freezes the first column of the worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <returns>The worksheet for method chaining.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        public static Worksheet FreezeFirstColumn(this Worksheet worksheet)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            worksheet.FreezePane(0, 1);
            return worksheet;
        }

        /// <summary>
        /// Freezes both the top row and first column of the worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <returns>The worksheet for method chaining.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        public static Worksheet FreezeTopRowAndFirstColumn(this Worksheet worksheet)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            worksheet.FreezePane(1, 1);
            return worksheet;
        }

        /// <summary>
        /// Freezes the specified number of top rows.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="rowCount">The number of rows to freeze.</param>
        /// <returns>The worksheet for method chaining.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when rowCount is negative.</exception>
        public static Worksheet FreezeTopRows(this Worksheet worksheet, int rowCount)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            if (rowCount < 0)
                throw new ArgumentOutOfRangeException(nameof(rowCount), "Row count cannot be negative.");

            worksheet.FreezePane(rowCount, 0);
            return worksheet;
        }

        /// <summary>
        /// Freezes the specified number of leftmost columns.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="columnCount">The number of columns to freeze.</param>
        /// <returns>The worksheet for method chaining.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when columnCount is negative.</exception>
        public static Worksheet FreezeLeftColumns(this Worksheet worksheet, int columnCount)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            if (columnCount < 0)
                throw new ArgumentOutOfRangeException(nameof(columnCount), "Column count cannot be negative.");

            worksheet.FreezePane(0, columnCount);
            return worksheet;
        }

        /// <summary>
        /// Freezes the panes at a specified cell position, freezing all rows above and all columns to the left of the cell.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="cellReference">The cell reference (e.g., "B3") indicating the position for the freeze pane.</param>
        /// <returns>The worksheet for method chaining.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet or cellReference is null.</exception>
        /// <exception cref="FormatException">Thrown when cellReference format is invalid.</exception>
        public static Worksheet FreezePanesAt(this Worksheet worksheet, string cellReference)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            if (string.IsNullOrEmpty(cellReference))
                throw new ArgumentNullException(nameof(cellReference));

            // Parse cell reference to get row and column
            var match = System.Text.RegularExpressions.Regex.Match(cellReference, @"([A-Z]+)(\d+)");
            if (!match.Success)
                throw new FormatException($"Invalid cell reference format: {cellReference}");

            string columnPart = match.Groups[1].Value;
            int rowNumber = int.Parse(match.Groups[2].Value);

            // Convert column letter to column number
            int columnNumber = 0;
            foreach (char c in columnPart)
            {
                columnNumber = columnNumber * 26 + (c - 'A' + 1);
            }

            // Freeze panes at the specified position
            worksheet.FreezePane(rowNumber - 1, columnNumber);
            return worksheet;
        }

        /// <summary>
        /// Removes all freeze panes from the worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <returns>The worksheet for method chaining.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        public static Worksheet UnfreezePanes(this Worksheet worksheet)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            worksheet.FreezePane(0, 0);
            return worksheet;
        }
    }
}