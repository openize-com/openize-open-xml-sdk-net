using System;
using DocumentFormat.OpenXml;
using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;

namespace Openize.Cells
{
    public sealed class Cell
    {

        private readonly DocumentFormat.OpenXml.Spreadsheet.Cell _cell;
        private readonly WorkbookPart _workbookPart;

        private readonly SheetData _sheetData;

        /// <summary>
        /// Gets the cell reference in A1 notation.
        /// </summary>
        public string CellReference => _cell.CellReference;

        /// <summary>
        /// Initializes a new instance of the <see cref="Cell"/> class.
        /// </summary>
        /// <param name="cell">The underlying OpenXML cell object.</param>
        /// <param name="sheetData">The sheet data containing the cell.</param>
        /// <exception cref="ArgumentNullException">
        /// Thrown when <paramref name="cell"/> or <paramref name="sheetData"/> is null.
        /// </exception>
        public Cell(DocumentFormat.OpenXml.Spreadsheet.Cell cell, SheetData sheetData, WorkbookPart workbookPart)
        {
            _cell = cell ?? throw new ArgumentNullException(nameof(cell));
            _sheetData = sheetData ?? throw new ArgumentNullException(nameof(sheetData));
            _workbookPart = workbookPart ?? throw new ArgumentNullException(nameof(workbookPart));
        }

        /// <summary>
        /// Adds a hyperlink to the specified cell with an optional tooltip.
        /// </summary>
        /// <param name="hyperlinkUrl">The URL of the hyperlink to be added.</param>
        /// <param name="tooltip">The optional tooltip to display when hovering over the hyperlink.</param>
        /// <exception cref="ArgumentException">
        /// Thrown when <paramref name="hyperlinkUrl"/> is null or empty.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the parent WorksheetPart or the OpenXML Worksheet cannot be found.
        /// </exception>
        /// <remarks>
        /// This method creates a hyperlink relationship in the parent worksheet and adds a Hyperlink element
        /// to the cell's reference in the worksheet. The method does not modify the cell's text; use <see cref="PutValue"/>
        /// to set display text if needed.
        /// </remarks>
        /// <example>
        /// The following example demonstrates how to use the <c>SetHyperlink</c> method:
        /// <code>
        /// using (Workbook wb = new Workbook("path/to/your/file.xlsx"))
        /// {
        ///     Worksheet sheet = wb.Worksheets[0];
        ///     Cell cell = sheet.Cells["A1"];
        ///     cell.PutValue("Click Me");
        ///     cell.SetHyperlink("https://example.com", "Visit Example");
        ///     wb.Save("path/to/your/file.xlsx");
        /// }
        /// </code>
        /// </example>
        public void SetHyperlink(string hyperlinkUrl, string tooltip = null)
        {
            if (string.IsNullOrEmpty(hyperlinkUrl))
                throw new ArgumentException("Hyperlink URL cannot be null or empty.", nameof(hyperlinkUrl));

            // Get the WorksheetPart for this cell
            WorksheetPart worksheetPart = GetWorksheetPart();
            if (worksheetPart == null)
                throw new InvalidOperationException("WorksheetPart is not available for this cell.");

            // Get the underlying OpenXML Worksheet
            var openXmlWorksheet = worksheetPart.Worksheet;
            if (openXmlWorksheet == null)
                throw new InvalidOperationException("The OpenXML Worksheet is not available.");

            // Ensure the Hyperlinks collection exists
            Hyperlinks hyperlinks = openXmlWorksheet.GetFirstChild<Hyperlinks>();
            if (hyperlinks == null)
            {
                hyperlinks = new Hyperlinks();
                openXmlWorksheet.InsertAfter(hyperlinks, openXmlWorksheet.GetFirstChild<SheetData>());
            }

            // Create a unique relationship ID
            string relationshipId = "rId" + Guid.NewGuid().ToString();

            // Add the hyperlink relationship to the worksheet part
            worksheetPart.AddHyperlinkRelationship(new Uri(hyperlinkUrl, UriKind.Absolute), true, relationshipId);

            // Create and add the Hyperlink element
            Hyperlink hyperlink = new Hyperlink
            {
                Reference = CellReference, // The reference of this cell in A1 notation
                Tooltip = tooltip,
                Id = relationshipId
            };
            hyperlinks.Append(hyperlink);

            // Save changes to the worksheet
            openXmlWorksheet.Save();
        }

        /// <summary>
        /// Retrieves the WorksheetPart containing this cell.
        /// </summary>
        /// <returns>The <see cref="WorksheetPart"/> that contains the current cell.</returns>
        /// <exception cref="InvalidOperationException">
        /// Thrown if the WorksheetPart cannot be located.
        /// </exception>
        /// <remarks>
        /// This method searches all WorksheetParts in the parent WorkbookPart and identifies
        /// the one containing the current SheetData.
        /// </remarks>
        private WorksheetPart GetWorksheetPart()
        {
            foreach (var worksheetPart in _workbookPart.GetPartsOfType<WorksheetPart>())
            {
                if (worksheetPart.Worksheet.Descendants<SheetData>().Contains(_sheetData))
                {
                    return worksheetPart;
                }
            }

            throw new InvalidOperationException("Unable to find the WorksheetPart containing this cell.");
        }

        /// <summary>
        /// Sets the value of the cell as a string.
        /// </summary>
        /// <param name="value">The value to set.</param>
        public void PutValue(string value)
        {
            PutValue(value, CellValues.String);
        }

        /// <summary>
        /// Sets the value of the cell as a number.
        /// </summary>
        /// <param name="value">The numeric value to set.</param>
        public void PutValue(double value)
        {
            PutValue(value.ToString(CultureInfo.InvariantCulture), CellValues.Number);
        }

        /// <summary>
        /// Sets the value of the cell as a date.
        /// </summary>
        /// <param name="value">The date value to set.</param>
        public void PutValue(DateTime value)
        {
            PutValue(value.ToOADate().ToString(CultureInfo.InvariantCulture), CellValues.Date);
        }

        /// <summary>
        /// Sets the cell's value with a specific data type.
        /// </summary>
        /// <param name="value">The value to set.</param>
        /// <param name="dataType">The data type of the value.</param>
        private void PutValue(string value, CellValues dataType)
        {
            _cell.DataType = new EnumValue<CellValues>(dataType);
            _cell.CellValue = new CellValue(value);

        }

        /// <summary>
        /// Sets a formula for the cell.
        /// </summary>
        /// <param name="formula">The formula to set.</param>
        public void PutFormula(string formula)
        {
            _cell.CellFormula = new CellFormula(formula);
            _cell.CellValue = new CellValue(); // You might want to set some default value or calculated value here
        }

        /// <summary>
        /// Gets the value of the cell.
        /// </summary>
        /// <returns>The cell value as a string.</returns>
        public string GetValue()
        {
            if (_cell == null || _cell.CellValue == null) return "";

            if (_cell.DataType != null && _cell.DataType.Value == CellValues.SharedString)
            {
                int index = int.Parse(_cell.CellValue.Text);
                SharedStringTablePart sharedStrings = _workbookPart.SharedStringTablePart;
                return sharedStrings.SharedStringTable.ElementAt(index).InnerText;
            }
            else
            {
                return _cell.CellValue.Text;
            }
        }

        /// <summary>
        /// Gets the data type of the cell's value.
        /// </summary>
        /// <returns>The cell's value data type, or null if not set.</returns>
        public CellValues? GetDataType()
        {
            return _cell.DataType?.Value;
        }


        /// <summary>
        /// Gets the formula set for the cell.
        /// </summary>
        /// <returns>The cell's formula as a string, or null if not set.</returns>
        public string GetFormula()
        {
            return _cell.CellFormula?.Text;
        }

        /// <summary>
        /// Applies a style to the cell.
        /// </summary>
        /// <param name="styleIndex">The index of the style to apply.</param>
        public void ApplyStyle(uint styleIndex)
        {
            _cell.StyleIndex = styleIndex;
        }
    }

}

