using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;


namespace Openize.Cells
{
    /// <summary>
    /// Represents a worksheet within an Excel file, providing methods to manipulate its content.
    /// </summary>
    public sealed class Worksheet
    {
        private WorksheetPart _worksheetPart;
        private SheetData _sheetData;
        private DocumentFormat.OpenXml.Spreadsheet.Cell sourceCell;
        public const double DefaultColumnWidth = 8.43; // Default width in character units
        public const double DefaultRowHeight = 15.0;   // Default height in points

        private WorkbookPart _workbookPart;
        /// <summary>
        /// Gets the CellIndexer for the worksheet. This property provides indexed access to the cells of the worksheet.
        /// </summary>
        /// <value>
        /// The CellIndexer for the worksheet.
        /// </value>
        public CellIndexer Cells { get; }

        /// <summary>
        /// Initializes a new instance of the Worksheet class with the specified WorksheetPart and Worksheet.
        /// This constructor sets up the internal structure of the Worksheet object, including initializing the Cells property.
        /// </summary>
        /// <param name="worksheetPart">The WorksheetPart associated with this Worksheet. Cannot be null.</param>
        /// <param name="worksheet">The OpenXml Spreadsheet.Worksheet object. Cannot be null.</param>
        /// <exception cref="ArgumentNullException">
        /// Thrown if the provided worksheetPart is null.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown if SheetData is not found in the provided worksheet.
        /// </exception>
        private Worksheet(WorksheetPart worksheetPart, DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet, WorkbookPart workbookPart)
        {
            _worksheetPart = worksheetPart ?? throw new ArgumentNullException(nameof(worksheetPart));

            _sheetData = worksheet?.Elements<SheetData>().FirstOrDefault()
                         ?? throw new InvalidOperationException("SheetData not found in the worksheet.");
            _workbookPart = workbookPart ?? throw new ArgumentNullException(nameof(workbookPart));

            // Initialize the Cells property
            this.Cells = new CellIndexer(this);

        }

        /// <summary>
        /// Creates an instance of the Worksheet class using the provided WorksheetPart and Worksheet.
        /// This method serves as a factory for creating Worksheet objects, encapsulating the instantiation logic.
        /// </summary>
        /// <param name="worksheetPart">The WorksheetPart to be associated with the Worksheet. Cannot be null.</param>
        /// <param name="worksheet">The OpenXml Spreadsheet.Worksheet object to be associated with the Worksheet. Cannot be null.</param>
        /// <returns>A new instance of the Worksheet class.</returns>
        public class WorksheetFactory
        {
            public static Worksheet CreateInstance(WorksheetPart worksheetPart, DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet, WorkbookPart workbookPart)
            {
                return new Worksheet(worksheetPart, worksheet, workbookPart);
            }
        }

        /// <summary>
        /// Gets or sets the name of the worksheet. This property performs several checks to ensure the integrity and validity of the worksheet and workbook parts.
        /// </summary>
        /// <value>
        /// The name of the worksheet.
        /// </value>
        /// <exception cref="InvalidOperationException">
        /// Thrown if the WorksheetPart is null, if the WorkbookPart is not found as a parent, if the ID of the part is null or empty, or if no sheet is found with the specified ID.
        /// </exception>
        /// <exception cref="ArgumentException">
        /// Thrown when attempting to set the name with a null or empty value.
        /// </exception>
        public string Name
        {
            get
            {
                if (_worksheetPart == null)
                    throw new InvalidOperationException("WorksheetPart is null.");

                var workbookPart = _worksheetPart.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
                if (workbookPart == null)
                    throw new InvalidOperationException("WorkbookPart not found as a parent.");

                var id = workbookPart.GetIdOfPart(_worksheetPart);
                if (string.IsNullOrEmpty(id))
                    throw new InvalidOperationException("ID is null or empty.");

                var sheet = workbookPart.Workbook.Sheets.Cast<Sheet>().FirstOrDefault(s => s.Id.Value.Equals(id));
                if (sheet == null)
                    throw new InvalidOperationException("Sheet not found with the specified ID.");

                return sheet.Name;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                    throw new ArgumentException("Sheet name cannot be null or empty", nameof(value));

                var workbookPart = _worksheetPart.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
                var sheet = workbookPart?.Workbook.Sheets.Cast<Sheet>()
                              .FirstOrDefault(s => s.Id.Value.Equals(workbookPart.GetIdOfPart(_worksheetPart)));

                if (sheet != null)
                    sheet.Name = value;
            }
        }

        /// <summary>
        /// Gets the row index where the pane is frozen in the worksheet.
        /// Returns 0 if no rows are frozen.
        /// </summary>
        /// <remarks>
        /// This property retrieves the value of <see cref="VerticalSplit"/> from the Pane element in the SheetView.
        /// </remarks>
        public int FreezePanesRow
        {
            get
            {
                var pane = _worksheetPart.Worksheet.GetFirstChild<SheetViews>()
                             ?.Elements<SheetView>().FirstOrDefault()
                             ?.Elements<Pane>().FirstOrDefault();
                return pane != null ? (int)pane.VerticalSplit.Value : 0;
            }
        }

        /// <summary>
        /// Gets the column index where the pane is frozen in the worksheet.
        /// Returns 0 if no columns are frozen.
        /// </summary>
        /// <remarks>
        /// This property retrieves the value of <see cref="HorizontalSplit"/> from the Pane element in the SheetView.
        /// </remarks>
        public int FreezePanesColumn
        {
            get
            {
                var pane = _worksheetPart.Worksheet.GetFirstChild<SheetViews>()
                             ?.Elements<SheetView>().FirstOrDefault()
                             ?.Elements<Pane>().FirstOrDefault();
                return pane != null ? (int)pane.HorizontalSplit.Value : 0;
            }
        }


        /// <summary>
        /// Retrieves a cell based on its reference.
        /// </summary>
        /// <param name="cellReference">The cell reference in A1 notation.</param>
        /// <returns>The cell at the specified reference.</returns>
        public Cell GetCell(string cellReference)
        {
            // This logic used to be in your indexer
            return new Cell(GetOrCreateCell(cellReference), _sheetData, _workbookPart);
        }

        /// <summary>
        /// Adds an image to the worksheet.
        /// </summary>
        /// <param name="image">The image to be added.</param>
        /// <param name="startRowIndex">The starting row index.</param>
        /// <param name="startColumnIndex">The starting column index.</param>
        /// <param name="endRowIndex">The ending row index.</param>
        /// <param name="endColumnIndex">The ending column index.</param>
        public void AddImage(Image image, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
            if (image == null) throw new ArgumentNullException(nameof(image));

            // Assuming you have a working constructor or factory method for ImageHandler
            var imgHandler = new ImageHandler(_worksheetPart);
            imgHandler.Add(image.Path, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
        }

        /// <summary>
        /// Extracts images from the worksheet.
        /// </summary>
        /// <returns>A list of images present in the worksheet.</returns>
        public List<Image> ExtractImages()
        {
            List<Image> imagePartsCollection = new List<Image>();

            if (this._worksheetPart.DrawingsPart == null)
                return imagePartsCollection; // Return an empty list instead of null

            foreach (var part in this._worksheetPart.DrawingsPart.ImageParts)
            {
                var stream = part.GetStream();
                var extension = GetImageExtension(part.ContentType);
                imagePartsCollection.Add(new Image(stream, extension));
            }
            return imagePartsCollection;
        }

        /// <summary>
        /// Sets the height of the specified row in the worksheet.
        /// </summary>
        /// <param name="rowIndex">The 1-based index of the row for which the height is to be set.</param>
        /// <param name="height">The height to set for the specified row.</param>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown if the rowIndex is less than 1 or if the height is less than 0.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown if the worksheet part or the worksheet is null.
        /// </exception>
        public void SetRowHeight(uint rowIndex, double height)
        {
            if (rowIndex < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index must be greater than 0.");
            }

            if (height < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(height), "Row height must be a positive number.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }
            var row = GetOrCreateRow(rowIndex);
            row.Height = height;
            row.CustomHeight = true;
        }

        /// <summary>
        /// Sets the width of the specified column in the worksheet.
        /// </summary>
        /// <param name="columnName">The name of the column (e.g., "A", "B", "C") for which the width is to be set.</param>
        /// <param name="width">The width to set for the specified column.</param>
        /// <exception cref="ArgumentNullException">
        /// Thrown if the columnName is null or empty.
        /// </exception>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown if the width is less than 0 or if the columnName is invalid or represents a column index out of range.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown if the worksheet part or the worksheet is null.
        /// </exception>
        public void SetColumnWidth(string columnName, double width)
        {
            if (string.IsNullOrWhiteSpace(columnName))
            {
                throw new ArgumentNullException(nameof(columnName), "Column name cannot be null or empty.");
            }

            if (width < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(width), "Column width must be a positive number.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }

            Columns columns = _worksheetPart.Worksheet.GetFirstChild<Columns>();
            if (columns == null)
            {
                columns = new Columns();
                _worksheetPart.Worksheet.InsertAfter(columns, _worksheetPart.Worksheet.GetFirstChild<SheetFormatProperties>());
            }

            uint columnIndex = (uint)ColumnLetterToIndex(columnName);
            var column = columns.Elements<Column>().FirstOrDefault(c => c.Min == columnIndex);
            if (column == null)
            {
                column = new Column { Min = columnIndex, Max = columnIndex, Width = width, CustomWidth = true };
                columns.Append(column);
            }
            else
            {
                column.Width = width;
                column.CustomWidth = true;
            }
        }

        /// <summary>
        /// Retrieves the width of the specified column in the worksheet.
        /// If the width of the column has been explicitly set, it returns that value; otherwise, it returns the default column width.
        /// </summary>
        /// <param name="columnIndex">The 1-based index of the column for which the width is to be retrieved.</param>
        /// <returns>
        /// The width of the specified column. If the column's width is explicitly set, that value is returned; 
        /// otherwise, the default column width is returned.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown if the columnIndex is less than 1.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown if the worksheet part or the worksheet is null.
        /// </exception>
        public double GetColumnWidth(uint columnIndex)
        {
            if (columnIndex < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(columnIndex), "Column index must be greater than 0.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }
            // Access the Columns collection
            var columns = _worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Columns>();
            if (columns != null)
            {
                foreach (var column in columns.Elements<Column>())
                {

                    // Explicitly cast Min and Max to uint and check for null
                    uint min = column.Min.HasValue ? column.Min.Value : uint.MinValue;
                    uint max = column.Max.HasValue ? column.Max.Value : uint.MaxValue;

                    if (columnIndex >= min && columnIndex <= max)
                    {
                        // Also check if Width is set
                        return column.Width.HasValue ? column.Width.Value : DefaultColumnWidth;
                    }
                }
            }

            return DefaultColumnWidth;
        }

        /// <summary>
        /// Auto-fits the specified column to match its content width by setting
        /// the <see cref="BestFit"/> and <see cref="CustomWidth"/> attributes.
        /// </summary>
        /// <param name="columnName">The column name (e.g., "A", "B") to auto-fit.</param>
        public void AutoFitColumn(string columnName)
        {
            var sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();

            Columns columns = _worksheetPart.Worksheet.Elements<Columns>().FirstOrDefault();
            if (columns == null)
            {
                columns = new Columns();
                _worksheetPart.Worksheet.InsertBefore(columns, sheetData);
            }

            int columnIndex = ConvertColumnNameToIndex(columnName);

            var column = new Column
            {
                Min = (uint)(columnIndex + 1),
                Max = (uint)(columnIndex + 1),
                Width = 26.83, // Optionally calculate width based on content
                BestFit = true,
                CustomWidth = true
            };
            columns.Append(column);

            _worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// Converts a column name (e.g., "A", "B", "AA") to a zero-based index.
        /// </summary>
        /// <param name="columnName">The Excel column name to convert.</param>
        /// <returns>The zero-based index of the specified column.</returns>
        private int ConvertColumnNameToIndex(string columnName)
        {
            int index = 0;
            foreach (char c in columnName.ToUpper())
            {
                index = index * 26 + (c - 'A' + 1);
            }
            return index - 1;
        }




        /// <summary>
        /// Retrieves the height of the specified row in the worksheet.
        /// If the height of the row has been explicitly set, it returns that value; otherwise, it returns the default row height.
        /// </summary>
        /// <param name="rowIndex">The 1-based index of the row for which the height is to be retrieved.</param>
        /// <returns>
        /// The height of the specified row. If the row's height is explicitly set, that value is returned; 
        /// otherwise, the default row height is returned.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown if the rowIndex is less than 1.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown if the worksheet part or the worksheet is null.
        /// </exception>
        public double GetRowHeight(uint rowIndex)
        {
            if (rowIndex < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index must be greater than 0.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }

            // Assuming _worksheetPart is the OpenXML WorksheetPart
            var rows = _worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>().Elements<Row>();

            foreach (var row in rows)
            {
                // Check if this is the row we are looking for
                if (row.RowIndex.Value == rowIndex)
                {
                    // If Height is set, return it, otherwise return default height
                    return row.Height.HasValue ? row.Height.Value : DefaultRowHeight;
                }
            }

            return DefaultRowHeight; // Return default height if no specific height is set
        }




        /// <summary>
        /// Protects the worksheet with the specified password. This method applies various protection settings to the sheet, 
        /// including locking objects, scenarios, auto-filters, pivot tables, and other elements from editing.
        /// </summary>
        /// <param name="password">The password used to protect the worksheet. The password is hashed before being applied.</param>
        /// <remarks>
        /// This method creates a new SheetProtection object with specific settings and applies it to the worksheet.
        /// If an existing SheetProtection element is present, it is removed before applying the new protection.
        /// After setting the protection, it saves the changes to the worksheet.
        /// </remarks>
        /// <exception cref="ArgumentException">
        /// Thrown if the provided <paramref name="password"/> is null or empty.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown if an error occurs while applying the protection, such as failing to save the changes.
        /// </exception
        public void ProtectSheet(string password)
        {
            SheetProtection sheetProtection = new SheetProtection()
            {
                Sheet = true,
                Objects = true,
                Scenarios = true,
                AutoFilter = true,
                PivotTables = true,
                Password = HashPassword(password),
                DeleteRows = true,
                DeleteColumns = true,
                FormatCells = true,
                FormatColumns = true,
                FormatRows = true,
                InsertColumns = true,
                InsertRows = true,
                InsertHyperlinks = true,
                Sort = true,
            };

            // Remove existing SheetProtection if any
            var existingProtection = _worksheetPart.Worksheet.Elements<SheetProtection>().FirstOrDefault();
            if (existingProtection != null)
            {
                existingProtection.Remove();
            }

            // Insert new SheetProtection after the SheetData element
            _worksheetPart.Worksheet.InsertAfter(sheetProtection, _worksheetPart.Worksheet.Elements<SheetData>().First());

            // Save the changes
            _worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// Determines whether the worksheet is protected.
        /// This method checks for the presence of a SheetProtection element within the worksheet to ascertain its protection status.
        /// </summary>
        /// <returns>
        /// A boolean value indicating whether the worksheet is protected. Returns <c>true</c> if the worksheet is protected, otherwise <c>false</c>.
        /// </returns>
        /// <remarks>
        /// A worksheet is considered protected if there is at least one SheetProtection element present in its elements.
        /// </remarks>
        public bool IsProtected()
        {
            return _worksheetPart.Worksheet.Elements<SheetProtection>().Any();
        }


        /// <summary>
        /// Removes protection from the worksheet, if it is currently protected.
        /// This method checks for the presence of a SheetProtection element and removes it to unprotect the sheet.
        /// </summary>
        /// <remarks>
        /// If the worksheet is protected (indicated by the presence of a SheetProtection element), this method removes the protection.
        /// After altering the protection status, it saves the changes to the worksheet.
        /// If the worksheet is not protected, this method performs no action.
        /// </remarks>
        /// <exception cref="InvalidOperationException">
        /// Thrown if an attempt is made to remove protection but no SheetProtection element is found. This should not normally occur, as the method first checks if the sheet is protected.
        /// </exception>
        public void UnprotectSheet()
        {
            if (IsProtected())
            {
                var sheetProtection = _worksheetPart.Worksheet.Elements<SheetProtection>().First();
                sheetProtection.Remove();

            }
            // Save the changes
            _worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// Hashes the given password.
        /// </summary>
        /// <param name="password">The password to hash.</param>
        /// <returns>The hashed password.</returns>
        private string HashPassword(string password)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(password));
        }



        /// <summary>
        /// Converts a column letter to its corresponding index.
        /// </summary>
        /// <param name="column">The column letter (e.g., "A", "B", ...).</param>
        /// <returns>The index corresponding to the column letter.</returns>
        private static int ColumnLetterToIndex(string column)
        {
            int index = 0;
            foreach (var ch in column)
            {
                index = (index * 26) + (ch - 'A' + 1);
            }
            return index;
        }


        /// <summary>
        /// Gets the file extension corresponding to a specific image content type.
        /// </summary>
        /// <param name="contentType">The image content type.</param>
        /// <returns>The file extension.</returns>
        private static string GetImageExtension(string contentType)
        {
            switch (contentType.ToLower())
            {
                case "image/jpeg": return "jpeg";
                case "image/png": return "png";
                case "image/gif": return "gif";
                case "image/tiff": return "tiff";
                case "image/bmp": return "bmp";
                default: throw new ArgumentOutOfRangeException(nameof(contentType), $"Unsupported image content type: {contentType}");
            }
        }


        /// <summary>
        /// Retrieves or creates a cell for a specific cell reference.
        /// </summary>
        /// <param name="cellReference">The cell reference in A1 notation.</param>
        /// <returns>The corresponding cell.</returns>
        private DocumentFormat.OpenXml.Spreadsheet.Cell GetOrCreateCell(string cellReference)
        {

            var cell = _sheetData.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>()
                                .FirstOrDefault(c => string.Equals(c.CellReference.Value, cellReference, StringComparison.OrdinalIgnoreCase));

            if (cell == null)
            {
                cell = new DocumentFormat.OpenXml.Spreadsheet.Cell { CellReference = cellReference };
                var rowIndex = GetRowIndex(cellReference);
                var row = GetOrCreateRow(rowIndex);
                row.Append(cell);
            }

            return cell;
        }

        /// <summary>
        /// Retrieves the row index from a cell reference.
        /// </summary>
        /// <param name="cellReference">The cell reference in A1 notation.</param>
        /// <returns>The row index.</returns>
        private uint GetRowIndex(string cellReference)
        {
            var match = Regex.Match(cellReference, @"\d+");
            if (!match.Success)
                throw new FormatException("Invalid cell reference format.");

            return uint.Parse(match.Value);
        }

        /// <summary>
        /// Retrieves or creates a row for a specific row index.
        /// </summary>
        /// <param name="rowIndex">The row index.</param>
        /// <returns>The corresponding row.</returns>
        private Row GetOrCreateRow(uint rowIndex)
        {
            var row = _sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                row = new Row { RowIndex = rowIndex };
                _sheetData.Append(row);
            }
            return row;
        }


        /// <summary>
        /// Merges a range of cells specified by the start and end cell references in A1 notation.
        /// This method creates a merged cell area that spans from the start cell to the end cell.
        /// </summary>
        /// <param name="startCellReference">The start cell reference in A1 notation for the merge range.</param>
        /// <param name="endCellReference">The end cell reference in A1 notation for the merge range.</param>
        /// <remarks>
        /// If a MergeCells element already exists in the worksheet, this method appends a new merge cell reference to it.
        /// If no MergeCells element exists, it creates a new one and then appends the merge cell reference.
        /// After defining the merge cell range, it saves the changes to the WorksheetPart.
        /// </remarks>
        /// <exception cref="ArgumentException">
        /// Thrown when either <paramref name="startCellReference"/> or <paramref name="endCellReference"/> is null, empty, or invalid.
        /// </exception>
        public void MergeCells(string startCellReference, string endCellReference)
        {
            if (_worksheetPart.Worksheet.Elements<MergeCells>().Any())
            {
                // MergeCells element already exists, use it
                MergeCells mergeCells = _worksheetPart.Worksheet.Elements<MergeCells>().First();
                MergeCell newMergeCell = new MergeCell() { Reference = new StringValue(startCellReference + ":" + endCellReference) };
                mergeCells.Append(newMergeCell);
            }
            else
            {
                // Otherwise, create new MergeCells element
                MergeCells mergeCells = new MergeCells();
                MergeCell newMergeCell = new MergeCell() { Reference = new StringValue(startCellReference + ":" + endCellReference) };
                mergeCells.Append(newMergeCell);
                _worksheetPart.Worksheet.InsertAfter(mergeCells, _worksheetPart.Worksheet.Elements<SheetData>().First());
            }

            // Save changes to the WorksheetPart
            _worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// Retrieves the index of the current worksheet within the workbook. This method locates the worksheet within the workbook's collection of sheets and returns its index.
        /// Note that the SheetId property of a worksheet is different from its index in the workbook's sheet collection.
        /// </summary>
        /// <returns>
        /// The index of the sheet within the workbook. This is not the same as the SheetId.
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// Thrown if no WorkbookPart is found, or if the worksheet is not found in the workbook.
        /// </exception>
        public int GetSheetIndex()
        {
            var workbookPart = _worksheetPart.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
            if (workbookPart == null)
                throw new InvalidOperationException("No WorkbookPart found.");

            var sheets = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>();
            var sheet = sheets.FirstOrDefault(s => workbookPart.GetPartById(s.Id) == _worksheetPart);

            if (sheet == null)
                throw new InvalidOperationException("Worksheet not found in workbook.");

            // Note: SheetId is not the same as the index of the sheet in the workbook.
            // If you specifically need the index, you may need to implement a different approach.
            return int.Parse(sheet.SheetId);
        }


        /// <summary>
        /// Retrieves a range of cells specified by the start and end row and column indices.
        /// </summary>
        /// <param name="startRowIndex">The starting row index of the range.</param>
        /// <param name="startColumnIndex">The starting column index of the range.</param>
        /// <param name="endRowIndex">The ending row index of the range.</param>
        /// <param name="endColumnIndex">The ending column index of the range.</param>
        /// <returns>
        /// A <see cref="Range"/> object representing the specified range of cells.
        /// </returns>
        public Range GetRange(uint startRowIndex, uint startColumnIndex, uint endRowIndex, uint endColumnIndex)
        {
            return new Range(this, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
        }

        /// <summary>
        /// Retrieves a range of cells specified by the start and end cell references in A1 notation.
        /// </summary>
        /// <param name="startCellReference">The start cell reference in A1 notation.</param>
        /// <param name="endCellReference">The end cell reference in A1 notation.</param>
        /// <returns>
        /// A <see cref="Range"/> object representing the specified range of cells.
        /// </returns>
        public Range GetRange(string startCellReference, string endCellReference)
        {
            var startCellParts = ParseCellReference(startCellReference);
            var endCellParts = ParseCellReference(endCellReference);
            return GetRange(startCellParts.row, startCellParts.column, endCellParts.row, endCellParts.column);
        }

        /// <summary>
        /// Adds a dropdown list validation to a specified cell. The dropdown list contains the options provided.
        /// </summary>
        /// <param name="cellReference">The cell reference in A1 notation where the dropdown should be added.</param>
        /// <param name="options">An array of string values that will appear as options in the dropdown list.</param>
        /// <exception cref="ArgumentException">
        /// Thrown when <paramref name="cellReference"/> is null or invalid, or if <paramref name="options"/> is empty or null.
        /// </exception>
        /// <remarks>
        /// This method creates a data validation rule that restricts input to the cell to the provided list of options.
        /// </remarks>
        public void AddDropdownListValidation(string cellReference, string[] options)
        {
            // Convert options array into a comma-separated string
            string formula = string.Join(",", options);

            // Create the data validation object
            DataValidation dataValidation = new DataValidation
            {
                Type = DataValidationValues.List,
                ShowDropDown = true,
                ShowErrorMessage = true,
                ErrorTitle = "Invalid input",
                Error = "The value entered is not in the list.",
                Formula1 = new Formula1("\"" + formula + "\""), // The formula is enclosed in quotes
                SequenceOfReferences = new ListValue<StringValue> { InnerText = cellReference }
            };

            // Add the data validation to the worksheet
            var dataValidations = _worksheetPart.Worksheet.GetFirstChild<DataValidations>();
            if (dataValidations == null)
            {
                dataValidations = new DataValidations();
                _worksheetPart.Worksheet.AppendChild(dataValidations);
            }

            dataValidations.AppendChild(dataValidation);
        }

        // <summary>
        /// Applies a data validation rule to a specific cell in the worksheet.
        /// </summary>
        /// <param name="cellReference">The reference of the cell to which the validation rule will be applied, e.g., "A1".</param>
        /// <param name="rule">The validation rule to apply to the cell.</param>
        /// <remarks>
        /// This method applies a specified data validation rule to a single cell in the worksheet. It first creates a <see cref="DataValidation"/> object based on the provided cell reference and validation rule, and then adds this data validation to the worksheet. This allows for dynamic application of validation criteria to cells, which is useful in scenarios where data integrity and input validation are required.
        /// </remarks>
        public void ApplyValidation(string cellReference, ValidationRule rule)
        {
            DataValidation dataValidation = CreateDataValidation(cellReference, rule);
            AddDataValidation(dataValidation);
        }

        /// <summary>
        /// Retrieves the validation rule applied to a specific cell in the worksheet.
        /// </summary>
        /// <param name="cellReference">The reference of the cell for which to retrieve the validation rule, e.g., "A1".</param>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the Worksheet part is not loaded or is null.
        /// </exception>
        /// <returns>
        /// The validation rule applied to the specified cell if one exists; otherwise, null.
        /// </returns>
        /// <remarks>
        /// This method searches for a data validation rule that applies to the specified cell. It iterates through all the data validation rules present in the worksheet. If a rule is found that includes the cell reference, the method constructs and returns a corresponding <see cref="ValidationRule"/> object. If no such rule is found, the method returns null. This is useful for dynamically determining validation criteria or rules applied to specific cells.
        /// </remarks>
        public ValidationRule GetValidationRule(string cellReference)
        {
            if (_worksheetPart == null)
            {
                throw new InvalidOperationException("Worksheet part is not loaded.");
            }

            var dataValidations = _worksheetPart.Worksheet.Descendants<DataValidation>();

            foreach (var dataValidation in dataValidations)
            {
                if (dataValidation.SequenceOfReferences.InnerText.Contains(cellReference))
                {
                    return CreateValidationRuleFromDataValidation(dataValidation);
                }
            }

            return null; // No validation rule found for the specified cell
        }

        private ValidationRule CreateValidationRuleFromDataValidation(DataValidation dataValidation)
        {
            // Determine the type of validation
            ValidationType type = DetermineValidationType(dataValidation.Type);

            // Depending on the type, create a corresponding ValidationRule
            switch (type)
            {
                case ValidationType.List:
                    // Assuming list validations use a comma-separated list of options in Formula1
                    var options = dataValidation.Formula1.Text.Split(',');
                    return new ValidationRule(options);

                case ValidationType.CustomFormula:
                    return new ValidationRule(dataValidation.Formula1.Text);

                // Add cases for other validation types (e.g., Date, WholeNumber, Decimal, TextLength)
                // ...

                default:
                    throw new NotImplementedException("Validation type not supported.");
            }
        }

        private ValidationType DetermineValidationType(DataValidationValues openXmlType)
        {
            // Map the OpenXML DataValidationValues to your ValidationType enum
            // This mapping depends on how closely your ValidationType enum aligns with OpenXML's types
            // Example mapping:
            
            switch (openXmlType.ToString())
            {
                case "List":
                    return ValidationType.List;
                case "Custom":
                    return ValidationType.CustomFormula;
                // Map other types...
                default:
                    throw new NotImplementedException("Validation type not supported.");
            }
            
            /**
            if (openXmlType == DataValidationValues.List)
                return ValidationType.List;
            else if (openXmlType == DataValidationValues.List)
                return ValidationType.CustomFormula;
            else
                throw new NotImplementedException("Validation type not supported.");
            **/
        }



        private DataValidation CreateDataValidation(string cellReference, ValidationRule rule)
        {
            DataValidation dataValidation = new DataValidation
            {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = cellReference },
                ShowErrorMessage = true,
                ErrorTitle = rule.ErrorTitle,
                Error = rule.ErrorMessage
            };

            switch (rule.Type)
            {
                case ValidationType.List:
                    dataValidation.Type = DataValidationValues.List;
                    dataValidation.Formula1 = new Formula1($"\"{string.Join(",", rule.Options)}\"");
                    break;

                case ValidationType.Date:
                    dataValidation.Type = DataValidationValues.Date;
                    dataValidation.Operator = DataValidationOperatorValues.Between;
                    dataValidation.Formula1 = new Formula1(rule.MinValue.HasValue ? rule.MinValue.Value.ToString() : "0");
                    dataValidation.Formula2 = new Formula2(rule.MaxValue.HasValue ? rule.MaxValue.Value.ToString() : "0");
                    break;

                case ValidationType.WholeNumber:
                    dataValidation.Type = DataValidationValues.Whole;
                    dataValidation.Operator = DataValidationOperatorValues.Between;
                    dataValidation.Formula1 = new Formula1(rule.MinValue.HasValue ? rule.MinValue.Value.ToString() : "0");
                    dataValidation.Formula2 = new Formula2(rule.MaxValue.HasValue ? rule.MaxValue.Value.ToString() : "0");
                    break;

                case ValidationType.Decimal:
                    dataValidation.Type = DataValidationValues.Decimal;
                    dataValidation.Operator = DataValidationOperatorValues.Between;
                    dataValidation.Formula1 = new Formula1(rule.MinValue.HasValue ? rule.MinValue.Value.ToString() : "0");
                    dataValidation.Formula2 = new Formula2(rule.MaxValue.HasValue ? rule.MaxValue.Value.ToString() : "0");
                    break;

                case ValidationType.TextLength:
                    dataValidation.Type = DataValidationValues.TextLength;
                    dataValidation.Operator = DataValidationOperatorValues.Between;
                    dataValidation.Formula1 = new Formula1(rule.MinValue.HasValue ? rule.MinValue.Value.ToString() : "0");
                    dataValidation.Formula2 = new Formula2(rule.MaxValue.HasValue ? rule.MaxValue.Value.ToString() : "0");
                    break;

                case ValidationType.CustomFormula:
                    dataValidation.Type = DataValidationValues.Custom;
                    dataValidation.Formula1 = new Formula1(rule.CustomFormula);
                    break;

                default:
                    throw new ArgumentException("Unsupported validation type.");
            }

            return dataValidation;
        }


        private void AddDataValidation(DataValidation dataValidation)
        {
            var dataValidations = _worksheetPart.Worksheet.GetFirstChild<DataValidations>();
            if (dataValidations == null)
            {
                dataValidations = new DataValidations();
                _worksheetPart.Worksheet.AppendChild(dataValidations);
            }
            dataValidations.AppendChild(dataValidation);
        }


        private (uint row, uint column) ParseCellReference(string cellReference)
        {
            var match = Regex.Match(cellReference, @"([A-Z]+)(\d+)");
            if (!match.Success)
                throw new FormatException("Invalid cell reference format.");

            uint row = uint.Parse(match.Groups[2].Value);
            uint column = (uint)ColumnLetterToIndex(match.Groups[1].Value);

            return (row, column);
        }

        /// <summary>
        /// Hides a specific column in the worksheet.
        /// </summary>
        /// <param name="columnName">The letter of the column to hide. Cannot be null or whitespace.</param>
        /// <exception cref="ArgumentNullException">
        /// Thrown when <paramref name="columnName"/> is null or whitespace, indicating that the column name cannot be null or empty.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the Worksheet or WorksheetPart is null.
        /// </exception>
        /// <remarks>
        /// This method hides a single column in the worksheet, specified by the column name. If the column does not already exist in the worksheet's column collection, it is created with the Hidden property set to true. If the column already exists, the Hidden property is set to true, effectively hiding the column. The method ensures that the specified column is hidden, whether it was previously defined or not.
        /// </remarks>
        public void HideColumn(string columnName)
        {
            if (string.IsNullOrWhiteSpace(columnName))
            {
                throw new ArgumentNullException(nameof(columnName), "Column name cannot be null or empty.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }

            Columns columns = _worksheetPart.Worksheet.GetFirstChild<Columns>();
            if (columns == null)
            {
                columns = new Columns();
                _worksheetPart.Worksheet.InsertAt(columns, 0);
            }

            uint columnIndex = (uint)ColumnLetterToIndex(columnName);
            Column column = columns.Elements<Column>().FirstOrDefault(c => c.Min == columnIndex);
            if (column == null)
            {
                // If the column doesn't exist, create it and set it as hidden
                column = new Column { Min = columnIndex, Max = columnIndex, Hidden = true };
                columns.Append(column);
            }
            else
            {
                // If the column exists, just set it as hidden
                column.Hidden = true;
            }

            // Save the changes to the worksheet part
            _worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// Unhides a specific column in the worksheet.
        /// </summary>
        /// <param name="columnName">The letter of the column to unhide. Cannot be null or whitespace.</param>
        /// <exception cref="ArgumentNullException">
        /// Thrown when <paramref name="columnName"/> is null or whitespace, as the column name cannot be null or empty.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the Worksheet or WorksheetPart is null.
        /// </exception>
        /// <remarks>
        /// This method unhides a single column specified by the column name. If the column does not exist in the worksheet's column collection, it is added with the Hidden property set to false. If the column exists, the Hidden property is set to false. This ensures that the specified column is effectively unhidden regardless of its initial state. The method checks for the existence of the column and adjusts the Hidden property accordingly.
        /// </remarks>
        public void UnhideColumn(string columnName)
        {
            if (string.IsNullOrWhiteSpace(columnName))
            {
                throw new ArgumentNullException(nameof(columnName), "Column name cannot be null or empty.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }

            Columns columns = _worksheetPart.Worksheet.Elements<Columns>().FirstOrDefault();
            if (columns == null)
            {
                columns = new Columns();
                _worksheetPart.Worksheet.InsertAt(columns, 0);
            }

            uint columnIndex = (uint)ColumnLetterToIndex(columnName);
            Column column = columns.Elements<Column>().FirstOrDefault(c => c.Min == columnIndex);

            if (column == null)
            {
                // If the column doesn't exist in the collection, add it and set Hidden to false
                column = new Column { Min = columnIndex, Max = columnIndex, Hidden = false };
                columns.Append(column);
            }
            else
            {
                // If the column exists, set Hidden to false or clear the attribute
                column.Hidden = false;
            }

            _worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// Unhides a range of columns in the worksheet.
        /// </summary>
        /// <param name="startColumn">The letter of the starting column to unhide. Cannot be null or empty.</param>
        /// <param name="numberOfColumns">The number of columns to unhide, starting from the startColumn. Must be greater than 0.</param>
        /// <exception cref="ArgumentException">
        /// Thrown when <paramref name="startColumn"/> is null or whitespace.
        /// </exception>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown when <paramref name="numberOfColumns"/> is less than or equal to 0, as the number of columns to unhide must be greater than 0.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the Worksheet or WorksheetPart is null.
        /// </exception>
        /// <remarks>
        /// This method unhides a specified range of columns in the worksheet. It calculates the column indices based on the starting column letter and the number of columns to unhide. If the columns within the specified range are not currently hidden or do not exist, no action is taken for those columns. The method only modifies columns that are defined and hidden.
        /// </remarks>
        public void UnhideColumns(string startColumn, int numberOfColumns)
        {
            if (string.IsNullOrWhiteSpace(startColumn))
            {
                throw new ArgumentException("Start column cannot be null or empty.", nameof(startColumn));
            }

            if (numberOfColumns <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(numberOfColumns), "Number of columns to unhide must be greater than 0.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }

            int startColumnIndex = ColumnLetterToIndex(startColumn);
            int endColumnIndex = startColumnIndex + numberOfColumns - 1;

            Columns columns = _worksheetPart.Worksheet.Elements<Columns>().FirstOrDefault();
            if (columns == null)
            {
                // If there are no columns defined, there is nothing to unhide
                return;
            }

            bool columnModified = false;
            foreach (Column column in columns)
            {
                if (column.Min <= endColumnIndex + 1 && column.Max >= startColumnIndex + 1)
                {
                    column.Hidden = null; // Unhide the column
                    columnModified = true;
                }
            }

            // Save the changes if any column was modified
            if (columnModified)
            {
                _worksheetPart.Worksheet.Save();
            }
        }


        public List<uint> GetHiddenRows()
        {
            List<uint> itemList = new List<uint>();

            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }


            // Retrieve hidden rows.
            itemList = _worksheetPart.Worksheet.Descendants<Row>()
                .Where((r) => r?.Hidden is not null && r.Hidden.Value)
                .Select(r => r.RowIndex?.Value)
                .Cast<uint>()
                .ToList();

            return itemList;
        }

        public List<uint> GetHiddenColumns()
        {
            List<uint> itemList = new List<uint>();

            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }


            // Retrieve hidden columns.
            var cols = _worksheetPart.Worksheet.Descendants<Column>().Where((c) => c?.Hidden is not null && c.Hidden.Value);

            foreach (Column item in cols)
            {
                if (item.Min is not null && item.Max is not null)
                {
                    for (uint i = item.Min.Value; i <= item.Max.Value; i++)
                    {
                        itemList.Add(i);
                    }
                }
            }

            return itemList;
        }

        /// <summary>
        /// Freezes the specified rows and/or columns of the worksheet.
        /// </summary>
        /// <param name="rowsToFreeze">The number of rows to freeze. Set to 0 if no row freezing is needed.</param>
        /// <param name="columnsToFreeze">The number of columns to freeze. Set to 0 if no column freezing is needed.</param>
        public void FreezePane(int rowsToFreeze, int columnsToFreeze)
        {
            // Ensure we are freezing at least one row or column
            if (rowsToFreeze == 0 && columnsToFreeze == 0)
            {
                return; // No freeze needed, exit the method.
            }

            // Retrieve or create the SheetViews element
            SheetViews sheetViews = _worksheetPart.Worksheet.GetFirstChild<SheetViews>();

            if (sheetViews == null)
            {
                sheetViews = new SheetViews();
                _worksheetPart.Worksheet.InsertAt(sheetViews, 0);
            }

            // Retrieve or create the SheetView element
            SheetView sheetView = sheetViews.Elements<SheetView>().FirstOrDefault();

            if (sheetView == null)
            {
                sheetView = new SheetView() { WorkbookViewId = (UInt32Value)0U };
                sheetViews.Append(sheetView);
            }

            // Remove any existing Pane elements to avoid conflicts
            Pane existingPane = sheetView.Elements<Pane>().FirstOrDefault();
            if (existingPane != null)
            {
                existingPane.Remove();
            }

            // Calculate the top left cell after the freeze
            string topLeftCell = GetTopLeftCell(rowsToFreeze, columnsToFreeze);

            // Define freeze pane settings dynamically based on the rows and columns to freeze
            Pane pane = new Pane
            {
                VerticalSplit = rowsToFreeze > 0 ? (double)rowsToFreeze : 0D,
                HorizontalSplit = columnsToFreeze > 0 ? (double)columnsToFreeze : 0D,
                TopLeftCell = topLeftCell,
                ActivePane = PaneValues.BottomRight,
                State = PaneStateValues.Frozen
            };

            // Adjust active pane based on what is being frozen
            if (rowsToFreeze > 0 && columnsToFreeze > 0)
            {
                pane.ActivePane = PaneValues.BottomRight; // Both rows and columns
            }
            else if (rowsToFreeze > 0)
            {
                pane.ActivePane = PaneValues.BottomLeft; // Only rows
            }
            else if (columnsToFreeze > 0)
            {
                pane.ActivePane = PaneValues.TopRight; // Only columns
            }

            // Insert the Pane as the first child of SheetView
            sheetView.InsertAt(pane, 0);

            // Add the selection for the frozen pane
            Selection selection = new Selection()
            {
                Pane = pane.ActivePane,
                ActiveCell = topLeftCell,
                SequenceOfReferences = new ListValue<StringValue>() { InnerText = topLeftCell }
            };

            // Ensure selection comes after the pane
            sheetView.Append(selection);

            // Save the worksheet
            _worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// Determines the top left cell after the freeze pane based on rows and columns to freeze.
        /// </summary>
        /// <param name="rowsToFreeze">The number of rows to freeze.</param>
        /// <param name="columnsToFreeze">The number of columns to freeze.</param>
        /// <returns>The top left cell reference as a string (e.g., "B2").</returns>
        private string GetTopLeftCell(int rowsToFreeze, int columnsToFreeze)
        {
            // Default top left cell is A1
            if (rowsToFreeze == 0 && columnsToFreeze == 0)
            {
                return "A1";
            }

            // Calculate column part (A, B, C, etc.) based on columns to freeze
            string columnLetter = columnsToFreeze > 0 ? GetColumnLetter(columnsToFreeze + 1) : "A";

            // Calculate row number based on rows to freeze
            int rowNumber = rowsToFreeze > 0 ? rowsToFreeze + 1 : 1;

            return $"{columnLetter}{rowNumber}";
        }

        /// <summary>
        /// Converts a column index (1-based) to an Excel column letter (A, B, C, ..., Z, AA, AB, etc.).
        /// </summary>
        /// <param name="columnIndex">The 1-based index of the column.</param>
        /// <returns>The corresponding column letter as a string.</returns>
        private string GetColumnLetter(int columnIndex)
        {
            string columnLetter = string.Empty;
            while (columnIndex > 0)
            {
                columnIndex--;
                columnLetter = (char)('A' + (columnIndex % 26)) + columnLetter;
                columnIndex /= 26;
            }
            return columnLetter;
        }

        /// <summary>
        /// Hides a specific row in the worksheet.
        /// </summary>
        /// <param name="rowIndex">The one-based index of the row to hide. Must be greater than 0.</param>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown when <paramref name="rowIndex"/> is 0, as the row index must be greater than 0.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the Worksheet or WorksheetPart is null, or if SheetData is null.
        /// </exception>
        /// <remarks>
        /// This method hides a single row specified by the rowIndex. If the row does not exist in the worksheet,
        /// it is created and then hidden. This ensures that the specified row is effectively hidden regardless of its
        /// initial existence. The method checks the existence of the row and sets the Hidden property accordingly.
        /// </remarks>
        public void HideRow(uint rowIndex)
        {
            if (rowIndex == 0)
            {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index must be greater than 0.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }

            SheetData sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                throw new InvalidOperationException("SheetData is null.");
            }

            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                // If the row doesn't exist, create it and set it as hidden
                row = new Row() { RowIndex = rowIndex, Hidden = true };
                sheetData.Append(row);
            }
            else
            {
                // If the row exists, just set it as hidden
                row.Hidden = true;
            }

            // Save the changes to the worksheet part
            _worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// Hides a range of rows in the worksheet.
        /// </summary>
        /// <param name="startRowIndex">The one-based index of the first row to hide. Must be greater than 0.</param>
        /// <param name="numberOfRows">The number of rows to hide, starting from the startRowIndex. Must be greater than 0.</param>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown when <paramref name="startRowIndex"/> or <paramref name="numberOfRows"/> is 0.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown when Worksheet or WorksheetPart is null, or if SheetData is null.
        /// </exception>
        /// <remarks>
        /// This method hides rows in a consecutive range starting from startRowIndex and spanning numberOfRows.
        /// If a row within the specified range does not exist, it is created and then hidden, ensuring that
        /// the entire specified range is effectively hidden. The method iterates through each row in the specified range
        /// and sets or creates the Hidden property as true.
        /// </remarks>
        public void HideRows(uint startRowIndex, uint numberOfRows)
        {
            if (startRowIndex == 0)
            {
                throw new ArgumentOutOfRangeException(nameof(startRowIndex), "Start row index must be greater than 0.");
            }

            if (numberOfRows == 0)
            {
                throw new ArgumentOutOfRangeException(nameof(numberOfRows), "Number of rows must be greater than 0.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }

            SheetData sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                throw new InvalidOperationException("SheetData is null.");
            }

            uint endRowIndex = startRowIndex + numberOfRows - 1;
            for (uint rowIndex = startRowIndex; rowIndex <= endRowIndex; rowIndex++)
            {
                Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
                if (row == null)
                {
                    // If the row doesn't exist, create it and set it as hidden
                    row = new Row() { RowIndex = rowIndex, Hidden = true };
                    sheetData.Append(row);
                }
                else
                {
                    // If the row exists, just set it as hidden
                    row.Hidden = true;
                }
            }

            // Save the changes to the worksheet part
            _worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// Unhides a single row in the worksheet.
        /// </summary>
        /// <param name="rowIndex">The one-based index of the row to unhide.</param>
        /// <remarks>
        /// This method unhides the row at the specified rowIndex. It is a convenience method that 
        /// internally calls <see cref="UnhideRows"/> with the numberOfRows parameter set to 1.
        /// If the row at the specified index does not exist or is already visible, 
        /// the method leaves it unaffected.
        /// </remarks>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown when the rowIndex is 0, as Excel row indices are 1-based.
        /// </exception>
        public void UnhideRow(uint rowIndex)
        {
            if (rowIndex == 0)
            {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index must be greater than 0.");
            }

            UnhideRows(rowIndex, 1);
        }

        /// <summary>
        /// Unhides a specified range of rows in the worksheet.
        /// </summary>
        /// <param name="startRowIndex">The one-based index of the first row to unhide.</param>
        /// <param name="numberOfRows">The number of rows to unhide, starting from the startRowIndex.</param>
        /// <remarks>
        /// This method unhides rows in a consecutive range starting from startRowIndex and covering numberOfRows. 
        /// If any row within the specified range does not exist or is already visible, the method leaves it unaffected.
        /// The method iterates through each row in the specified range and sets its Hidden property to false.
        /// </remarks>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown when the startRowIndex is 0 (since Excel row indices are 1-based) or 
        /// when numberOfRows is 0 (as at least one row must be specified to unhide).
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown when either the Worksheet or WorksheetPart is null, indicating that the worksheet has not been properly initialized,
        /// or when SheetData is null, indicating that the worksheet does not contain any rows.
        /// </exception>
        public void UnhideRows(uint startRowIndex, uint numberOfRows)
        {
            if (startRowIndex == 0)
            {
                throw new ArgumentOutOfRangeException(nameof(startRowIndex), "Start row index must be greater than 0.");
            }

            if (numberOfRows == 0)
            {
                throw new ArgumentOutOfRangeException(nameof(numberOfRows), "Number of rows must be greater than 0.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }

            SheetData sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                throw new InvalidOperationException("SheetData is null.");
            }

            uint endRowIndex = startRowIndex + numberOfRows - 1;
            for (uint rowIndex = startRowIndex; rowIndex <= endRowIndex; rowIndex++)
            {
                Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
                if (row != null)
                {
                    // Only unhide the row if it exists
                    row.Hidden = false;
                }
            }

            // Save the changes to the worksheet part
            _worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// Inserts a new row into the worksheet at the specified row index.
        /// </summary>
        /// <param name="rowIndex">The one-based index at which to insert the new row. Existing rows starting from this index will be shifted down.</param>
        /// <remarks>
        /// The method will shift down all existing rows and their cells starting from the specified row index. 
        /// Each cell's reference in the shifted rows will also be updated to reflect the new row index. 
        /// If a row already exists at the specified index, it will be shifted down along with subsequent rows.
        /// </remarks>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when the provided row index is 0, as row indices in Excel are 1-based.</exception>
        /// <exception cref="InvalidOperationException">Thrown when the Worksheet or WorksheetPart is null, or if the SheetData is not available in the Worksheet.</exception>
        public void InsertRow(uint rowIndex)
        {
            if (rowIndex == 0)
            {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index must be greater than 0.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null || _worksheetPart.Worksheet.GetFirstChild<SheetData>() == null)
            {
                throw new InvalidOperationException("Worksheet, WorksheetPart or SheetData is null.");
            }

            SheetData sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();

            // Shift the existing rows and their cells down by one
            var rowsToShift = sheetData.Elements<Row>().Where(r => r.RowIndex.Value >= rowIndex).ToList();
            foreach (Row row in rowsToShift)
            {
                row.RowIndex.Value++;

                foreach (DocumentFormat.OpenXml.Spreadsheet.Cell openXmlCell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                {
                    string newCellReference = IncrementCellReference(openXmlCell.CellReference, 1);
                    openXmlCell.CellReference = new StringValue(newCellReference);
                }
            }

            // Insert the new row at the specified position
            Row newRow = new Row() { RowIndex = rowIndex };
            sheetData.InsertBefore(newRow, sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex.Value > rowIndex));

            // Save the changes to the worksheet part
            _worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// Inserts a specified number of new rows into the worksheet starting at a given row index.
        /// </summary>
        /// <param name="startRowIndex">The one-based index of the row from which new rows should start being inserted. Must be greater than 0.</param>
        /// <param name="numberOfRows">The number of rows to insert. Must be greater than 0.</param>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown when either <paramref name="startRowIndex"/> or <paramref name="numberOfRows"/> is 0, as both values must be greater than 0.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the Worksheet, WorksheetPart, or SheetData is null.
        /// </exception>
        /// <remarks>
        /// This method inserts a number of new rows into the worksheet starting at the specified startRowIndex. Existing rows starting from this index are shifted downwards to make space for the new rows. This includes adjusting the row indices and references of existing cells to maintain data integrity. The method is useful in scenarios where rows need to be dynamically added to the worksheet without overwriting existing data.
        /// </remarks>
        public void InsertRows(uint startRowIndex, uint numberOfRows)
        {
            if (startRowIndex == 0)
            {
                throw new ArgumentOutOfRangeException(nameof(startRowIndex), "Start row index must be greater than 0.");
            }

            if (numberOfRows == 0)
            {
                throw new ArgumentOutOfRangeException(nameof(numberOfRows), "Number of rows to insert must be greater than 0.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null || _worksheetPart.Worksheet.GetFirstChild<SheetData>() == null)
            {
                throw new InvalidOperationException("Worksheet, WorksheetPart, or SheetData is null.");
            }

            SheetData sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();

            // Shift existing rows and their cells down by the number of rows
            var rowsToShift = sheetData.Elements<Row>().Where(r => r.RowIndex.Value >= startRowIndex).ToList();
            foreach (Row row in rowsToShift)
            {
                row.RowIndex.Value += numberOfRows;

                foreach (DocumentFormat.OpenXml.Spreadsheet.Cell openXmlCell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                {
                    string newCellReference = IncrementCellReference(openXmlCell.CellReference, (int)numberOfRows);
                    openXmlCell.CellReference = new StringValue(newCellReference);
                }
            }

            // Insert the new rows
            for (uint i = 0; i < numberOfRows; i++)
            {
                Row newRow = new Row() { RowIndex = startRowIndex + i };
                sheetData.InsertBefore(newRow, sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex.Value > startRowIndex + i));
            }

            // Save the changes to the worksheet part
            _worksheetPart.Worksheet.Save();
        }
        /// <summary>
        /// Retrieves the total number of rows in the worksheet.
        /// </summary>
        /// <returns>
        /// The total number of rows in the worksheet. Returns 0 if the worksheet or sheet data is null.
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the Worksheet or WorksheetPart is null.
        /// </exception>
        /// <remarks>
        /// This method calculates and returns the total number of rows present in the worksheet. It does this by counting the number of <see cref="Row"/> elements within the <see cref="SheetData"/>. If the worksheet or sheet data is null, indicating an improperly initialized or corrupted worksheet, the method returns 0. This is useful for dynamically determining the size of the worksheet and iterating through its rows.
        /// </remarks>
        public int GetRowCount()
        {
            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }

            SheetData sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                return 0; // No rows if SheetData is null
            }

            return sheetData.Elements<Row>().Count();
        }

        /// <summary>
        /// Retrieves the total number of columns in the worksheet.
        /// </summary>
        /// <returns>
        /// The total number of columns in the worksheet. Returns 0 if the worksheet or sheet data is null.
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the Worksheet or WorksheetPart is null.
        /// </exception>
        /// <remarks>
        /// This method calculates and returns the total number of columns present in the worksheet. It does this by analyzing the cell references in each row to determine the unique column indices in use. If the worksheet or sheet data is null, indicating an improperly initialized or corrupted worksheet, the method returns 0. This is useful for dynamically determining the size of the worksheet and iterating through its columns.
        /// </remarks>
        public int GetColumnCount()
        {
            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }

            SheetData sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                return 0; // No columns if SheetData is null
            }

            // HashSet to keep track of unique column indices
            var columnIndices = new HashSet<int>();

            foreach (Row row in sheetData.Elements<Row>())
            {
                foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                {
                    string cellReference = cell.CellReference;
                    if (!string.IsNullOrEmpty(cellReference))
                    {
                        // Extract the column part of the cell reference and convert it to an index
                        string columnPart = new String(cellReference.TakeWhile(Char.IsLetter).ToArray());
                        int columnIndex = ColumnLetterToIndex(columnPart);
                        columnIndices.Add(columnIndex);
                    }
                }
            }

            return columnIndices.Count;
        }

        /// <summary>
        /// Checks if a specific row in the worksheet is hidden.
        /// </summary>
        /// <param name="rowIndex">The one-based index of the row to check for hidden status.</param>
        /// <returns>
        /// <c>true</c> if the specified row is hidden; otherwise, <c>false</c>. Returns <c>false</c> if the worksheet or sheet data is null, or if the row doesn't exist.
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the Worksheet or WorksheetPart is null.
        /// </exception>
        /// <remarks>
        /// This method determines whether a specified row in the worksheet is hidden. It checks the existence of the row and its Hidden property. If the worksheet or sheet data is null, or if the row doesn't exist, the method returns <c>false</c>, indicating that the row is not hidden. This is useful for checking the visibility status of rows in the worksheet.
        /// </remarks>
        public bool IsRowHidden(uint rowIndex)
        {
            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }

            SheetData sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                // If there's no SheetData, the row is not hidden because it doesn't exist
                return false;
            }

            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                // If the row doesn't exist, it's not hidden
                return false;
            }

            // Check the Hidden property of the Row
            // If the row doesn't have the Hidden attribute, it's not hidden
            return row.Hidden != null && row.Hidden.Value;
        }

        /// <summary>
        /// Checks if a specific column in the worksheet is hidden.
        /// </summary>
        /// <param name="columnName">The letter of the column to check for hidden status. Cannot be null or whitespace.</param>
        /// <returns>
        /// <c>true</c> if the specified column is hidden; otherwise, <c>false</c>. Returns <c>false</c> if the worksheet or sheet data is null, or if the column doesn't exist.
        /// </returns>
        /// <exception cref="ArgumentNullException">
        /// Thrown when <paramref name="columnName"/> is null or whitespace, indicating that the column name cannot be null or empty.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the Worksheet or WorksheetPart is null.
        /// </exception>
        /// <remarks>
        /// This method determines whether a specified column in the worksheet is hidden. It checks the existence of the column in the worksheet's Columns collection and its Hidden property. If the worksheet or Columns collection is null, or if the column doesn't exist, the method returns <c>false</c>, indicating that the column is not hidden. This is useful for checking the visibility status of columns in the worksheet.
        /// </remarks>
        public bool IsColumnHidden(string columnName)
        {
            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }

            Columns columns = _worksheetPart.Worksheet.Elements<Columns>().FirstOrDefault();
            if (columns == null)
            {
                // If there are no Columns defined, the column is not hidden
                return false;
            }

            uint columnIndex = (uint)ColumnLetterToIndex(columnName);
            Column column = columns.Elements<Column>().FirstOrDefault(c => c.Min <= columnIndex && c.Max >= columnIndex);

            if (column == null)
            {
                // If the column doesn't exist in the collection, it's not hidden
                return false;
            }

            // Check the Hidden property of the Column
            return column.Hidden != null && column.Hidden.Value;
        }


        private static string IncrementCellReference(string reference, int rowCount)
        {
            var regex = new System.Text.RegularExpressions.Regex("([A-Za-z]+)(\\d+)");
            var match = regex.Match(reference);

            if (!match.Success) return reference;

            string columnReference = match.Groups[1].Value;
            int rowNumber = int.Parse(match.Groups[2].Value);

            return $"{columnReference}{rowNumber + rowCount}";
        }

        /// <summary>
        /// Inserts a new column to the right of a specified starting column.
        /// </summary>
        /// <param name="startColumn">The letter of the starting column for insertion. Cannot be null or whitespace.</param>
        /// <exception cref="ArgumentException">
        /// Thrown when <paramref name="startColumn"/> is null or whitespace, indicating that the starting column name cannot be null or empty.
        /// </exception>
        /// <remarks>
        /// This method inserts a new column to the right of the specified starting column in the worksheet. It shifts existing columns to the right to make space for the new column. All cell references in each row to the right of the starting column are adjusted accordingly to maintain data integrity. This is useful for dynamically adding columns to the worksheet without overwriting existing data.
        /// </remarks>
        public void InsertColumn(string startColumn)
        {
            InsertColumns(startColumn, 1);
        }

        /// <summary>
        /// Inserts a specified number of new columns to the right of a specified starting column.
        /// </summary>
        /// <param name="startColumn">The letter of the starting column for insertion. Cannot be null or whitespace.</param>
        /// <param name="numberOfColumns">The number of columns to insert. Must be greater than 0.</param>
        /// <exception cref="ArgumentException">
        /// Thrown when <paramref name="startColumn"/> is null or whitespace, indicating that the starting column name cannot be null or empty.
        /// </exception>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown when <paramref name="numberOfColumns"/> is less than or equal to 0, indicating that the number of columns to insert must be greater than 0.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the Worksheet, WorksheetPart, or SheetData is null.
        /// </exception>
        /// <remarks>
        /// This method inserts a specified number of new columns to the right of the specified starting column in the worksheet. It shifts existing columns to the right to make space for the new columns. All cell references in each row to the right of the starting column are adjusted accordingly to maintain data integrity. This is useful for dynamically adding columns to the worksheet without overwriting existing data.
        /// </remarks>
        public void InsertColumns(string startColumn, int numberOfColumns)
        {
            if (string.IsNullOrWhiteSpace(startColumn))
            {
                throw new ArgumentException("Start column cannot be null or empty.", nameof(startColumn));
            }

            if (numberOfColumns <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(numberOfColumns), "Number of columns to insert must be greater than 0.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null || _worksheetPart.Worksheet.GetFirstChild<SheetData>() == null)
            {
                throw new InvalidOperationException("Worksheet, WorksheetPart, or SheetData is null.");
            }

            int startColumnIndex = ColumnLetterToIndex(startColumn);
            SheetData sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();

            foreach (Row row in sheetData.Elements<Row>())
            {
                // Shift cell references in each row
                var cells = row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().ToList();
                foreach (var cell in cells)
                {
                    string cellReference = cell.CellReference;
                    int columnIndex = ColumnLetterToIndex(Regex.Match(cellReference, "[A-Za-z]+").Value);
                    if (columnIndex >= startColumnIndex)
                    {
                        string newCellReference = IncrementColumnReference(cellReference, numberOfColumns);
                        cell.CellReference = new StringValue(newCellReference);
                    }
                }
            }

            // Save the changes to the worksheet part
            _worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// Gets the column heading for a specified cell in the worksheet.
        /// </summary>
        /// <param name="cellName">The name of the cell (e.g., "A1").</param>
        /// <returns>The text of the column heading, or null if the column does not exist.</returns>
        /// <exception cref="ArgumentException">Thrown when the cellName is null or empty.</exception>
        /// <exception cref="InvalidOperationException">Thrown when the WorkbookPart is not found, the sheet is not found,
        /// the column name is invalid, no header cell is found, or the SharedStringTablePart is missing.</exception>
        /// <exception cref="IndexOutOfRangeException">Thrown when the shared string index is out of range.</exception>
        public string? GetColumnHeading(string cellName)
        {
            if (string.IsNullOrEmpty(cellName))
                throw new ArgumentException("Cell name cannot be null or empty.", nameof(cellName));

            var workbookPart = _worksheetPart.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
            if (workbookPart == null)
                throw new InvalidOperationException("No WorkbookPart found.");

            var sheets = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>();
            var sheet = sheets.FirstOrDefault(s => workbookPart.GetPartById(s.Id) == _worksheetPart);

            if (sheet == null)
                throw new InvalidOperationException("No matching sheet found for the provided WorksheetPart.");

            WorksheetPart worksheetPart = _worksheetPart;

            // Get the column name for the specified cell.
            string columnName = GetColumnName(cellName);

            if (string.IsNullOrEmpty(columnName))
                throw new InvalidOperationException("Unable to determine the column name from the provided cell name.");

            // Get the cells in the specified column and order them by row.
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Cell> cells = worksheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>()
                .Where(c => string.Compare(GetColumnName(c.CellReference?.Value), columnName, true) == 0)
                .OrderBy(r => GetRowIndexN(r.CellReference) ?? 0);

            if (!cells.Any())
            {
                // The specified column does not exist.
                return null;
            }

            // Get the first cell in the column.
            DocumentFormat.OpenXml.Spreadsheet.Cell headCell = cells.First();

            if (headCell == null)
                throw new InvalidOperationException("No header cell found in the specified column.");

            // If the content of the first cell is stored as a shared string, get the text of the first cell
            // from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
            if (headCell.DataType != null && headCell.DataType.Value == CellValues.SharedString && int.TryParse(headCell.CellValue?.Text, out int index))
            {
                var sharedStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sharedStringPart == null)
                    throw new InvalidOperationException("No SharedStringTablePart found.");

                var items = sharedStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();
                if (index < 0 || index >= items.Length)
                    throw new IndexOutOfRangeException("Shared string index is out of range.");

                return items[index].InnerText;
            }
            else
            {
                return headCell.CellValue?.Text;
            }
        }


        /// <summary>
        /// Gets the row index from the specified cell name.
        /// </summary>
        /// <param name="cellName">The cell name in A1 notation (e.g., "A1").</param>
        /// <returns>The row index as a nullable unsigned integer, or null if the cell name is invalid.</returns>
        /// <exception cref="FormatException">Thrown when the row index portion of the cell name cannot be parsed.</exception>
        private uint? GetRowIndexN(string? cellName)
        {
            if (cellName is null)
            {
                return null;
            }

            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }

        /// <summary>
        /// Given a cell name, parses the specified cell to get the column name.
        /// </summary>
        /// <param name="cellName">The cell name in A1 notation (e.g., "A1").</param>
        /// <returns>The column name as a string, or an empty string if the cell name is invalid.</returns>
        private string GetColumnName(string? cellName)
        {
            if (cellName is null)
            {
                return string.Empty;
            }

            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }
        private static string IncrementColumnReference(string reference, int columnCount)
        {
            var regex = new System.Text.RegularExpressions.Regex("([A-Za-z]+)(\\d+)");
            var match = regex.Match(reference);

            if (!match.Success) return reference;

            string columnLetters = match.Groups[1].Value;
            int rowNumber = int.Parse(match.Groups[2].Value);

            int columnIndex = ColumnLetterToIndex(columnLetters);
            int newColumnIndex = columnIndex + columnCount;
            string newColumnLetters = IndexToColumnLetter(newColumnIndex);

            return $"{newColumnLetters}{rowNumber}";
        }

        private static string IndexToColumnLetter(int index)
        {
            index++; // Adjust for 1-based index
            string columnLetter = string.Empty;
            while (index > 0)
            {
                int modulo = (index - 1) % 26;
                columnLetter = Convert.ToChar('A' + modulo) + columnLetter;
                index = (index - modulo) / 26;
            }
            return columnLetter;
        }


        /// <summary>
        /// Adds or updates a comment in a specified cell within the worksheet. If the cell already has a comment,
        /// it updates the existing comment text. If there is no comment, it creates a new one.
        /// </summary>
        /// <param name="cellReference">The cell reference where the comment should be added, e.g., "A1".</param>
        /// <param name="comment">The comment object containing the author and the text of the comment.</param>
        /// <remarks>
        /// This method ensures that the worksheet comments part exists before adding or updating a comment.
        /// It also manages the authors list to ensure that each author is only added once and reuses the existing author index if available.
        /// Usage of this method requires that the workbook and worksheet are properly initialized and that the worksheet part is correctly associated.
        /// </remarks>

        public void AddComment(string cellReference, Comment comment)
        {
            // Ensure the comments part exists
            var commentsPart = _worksheetPart.GetPartsOfType<WorksheetCommentsPart>().FirstOrDefault();
            CommentList commentList;
            Authors authors;

            if (commentsPart == null)
            {
                commentsPart = _worksheetPart.AddNewPart<WorksheetCommentsPart>();
                commentsPart.Comments = new Comments();

                // Initialize new CommentList and Authors only if a new comments part is created
                commentList = new CommentList();
                authors = new Authors();
                commentsPart.Comments.AppendChild(commentList);
                commentsPart.Comments.AppendChild(authors);
            }
            else
            {
                // Retrieve existing CommentList and Authors
                commentList = commentsPart.Comments.Elements<CommentList>().First();
                authors = commentsPart.Comments.Elements<Authors>().First();
            }

            // Ensure the author exists
            var author = authors.Elements<Author>().FirstOrDefault(a => a.Text == comment.Author);
            if (author == null)
            {
                author = new Author() { Text = comment.Author };
                authors.AppendChild(author);  // Use AppendChild to add to the XML structure
            }
            uint authorId = (uint)authors.Elements<Author>().ToList().IndexOf(author);

            // Add or update the comment
            var existingComment = commentList.Elements<DocumentFormat.OpenXml.Spreadsheet.Comment>().FirstOrDefault(c => c.Reference == cellReference);
            if (existingComment == null)
            {
                var newComment = new DocumentFormat.OpenXml.Spreadsheet.Comment() { Reference = cellReference, AuthorId = authorId };
                newComment.AppendChild(new CommentText(new DocumentFormat.OpenXml.Spreadsheet.Text(comment.Text)));
                commentList.AppendChild(newComment); // Ensure appending to commentList
            }
            else
            {
                // Update the existing comment's text
                existingComment.Elements<CommentText>().First().Text = new DocumentFormat.OpenXml.Spreadsheet.Text(comment.Text);
            }

            // Save the changes
            commentsPart.Comments.Save();
            _worksheetPart.Worksheet.Save();
        }



        public void CopyRange(Range sourceRange, string targetStartCellReference)
        {
            var (targetStartRow, targetStartColumn) = ParseCellReference(targetStartCellReference);

            uint rowOffset = targetStartRow - sourceRange.StartRowIndex;
            uint columnOffset = targetStartColumn - sourceRange.StartColumnIndex;

            for (uint row = sourceRange.StartRowIndex; row <= sourceRange.EndRowIndex; row++)
            {
                for (uint column = sourceRange.StartColumnIndex - 1; column < sourceRange.EndColumnIndex; column++)
                {
                    // Calculate target cell's row and column indices
                    uint targetRow = row + rowOffset;
                    uint targetColumn = column + columnOffset;

                    // Construct source and target cell references
                    string sourceCellRef = $"{IndexToColumnLetter((int)column)}{row}";
                    string targetCellRef = $"{IndexToColumnLetter((int)targetColumn)}{targetRow}";

                    this.Cells[targetCellRef].PutValue(this.Cells[sourceCellRef].GetValue());
                }
            }

            // Save the worksheet to apply changes
            _worksheetPart.Worksheet.Save();
        }




    }

    public class CellIndexer
    {
        private readonly Worksheet _worksheet;

        /// <summary>
        /// Initializes a new instance of the <see cref="CellIndexer"/> class.
        /// </summary>
        /// <param name="worksheet">The worksheet to index.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="worksheet"/> is null.</exception>
        public CellIndexer(Worksheet worksheet)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        }

        /// <summary>
        /// Gets the cell at the specified reference in A1 notation. This indexer provides a convenient way to access cells within the worksheet.
        /// </summary>
        /// <param name="cellReference">The cell reference in A1 notation.</param>
        /// <returns>The cell at the specified reference.</returns>
        public Cell this[string cellReference]
        {
            get
            {
                // Delegate the actual work to Worksheet class
                return _worksheet.GetCell(cellReference);
            }
        }

    }
}

