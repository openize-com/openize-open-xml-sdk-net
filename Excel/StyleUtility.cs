using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

/// <summary>
/// Specifies the horizontal alignment options for text within a cell.
/// </summary>
/// <remarks>
/// This enumeration is used to define how text is aligned horizontally within a cell.
/// The values map to corresponding Open XML horizontal alignment settings.
/// </remarks>
public enum HorizontalAlignment
{
    /// <summary>
    /// Aligns the text to the left edge of the cell.
    /// </summary>
    Left,

    /// <summary>
    /// Centers the text horizontally within the cell.
    /// </summary>
    Center,

    /// <summary>
    /// Aligns the text to the right edge of the cell.
    /// </summary>
    Right,

    /// <summary>
    /// Justifies the text so that it is evenly distributed across the width of the cell.
    /// </summary>
    Justify,

    /// <summary>
    /// Distributes the text evenly, including additional spacing between characters if necessary.
    /// </summary>
    Distributed
}

/// <summary>
/// Specifies the vertical alignment options for text within a cell.
/// </summary>
/// <remarks>
/// This enumeration is used to define how text is aligned vertically within a cell.
/// The values map to corresponding Open XML vertical alignment settings.
/// </remarks>
public enum VerticalAlignment
{
    /// <summary>
    /// Aligns the text to the top edge of the cell.
    /// </summary>
    Top,

    /// <summary>
    /// Centers the text vertically within the cell.
    /// </summary>
    Center,

    /// <summary>
    /// Aligns the text to the bottom edge of the cell.
    /// </summary>
    Bottom,

    /// <summary>
    /// Justifies the text so that it is evenly distributed across the height of the cell.
    /// </summary>
    Justify,

    /// <summary>
    /// Distributes the text evenly, including additional spacing between lines if necessary.
    /// </summary>
    Distributed
}

/// <summary>
/// Provides utility methods for creating and managing styles in an Excel workbook.
/// </summary>
/// <remarks>
/// This class facilitates the creation of custom styles, including font settings, colors, borders, 
/// and alignment, and integrates them into the workbook's stylesheet.
/// </remarks>
public class StyleUtility
{
    private readonly WorkbookStylesPart stylesPart;
    private readonly Stylesheet stylesheet;

    public uint CreateDefaultStyle()
    {
        return CreateStyle("Arial", 12, "000000"); 
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="StyleUtility"/> class with the specified workbook styles part.
    /// </summary>
    /// <param name="stylesPart">
    /// The <see cref="WorkbookStylesPart"/> instance representing the workbook's styles.
    /// This cannot be null.
    /// </param>
    /// <exception cref="ArgumentNullException">
    /// Thrown if <paramref name="stylesPart"/> is null.
    /// </exception>
    /// <remarks>
    /// During initialization, this constructor ensures that the workbook's stylesheet 
    /// is properly initialized with default structures for fonts, fills, borders, and cell formats.
    /// </remarks>
    public StyleUtility(WorkbookStylesPart stylesPart)
    {
        this.stylesPart = stylesPart ?? throw new ArgumentNullException(nameof(stylesPart));
        this.stylesheet = this.stylesPart.Stylesheet ?? new Stylesheet();

        // Initialize Fonts, Fills, Borders, and CellFormats if they are null
        this.stylesheet.Fonts = this.stylesheet.Fonts ?? new Fonts();
        this.stylesheet.Fills = this.stylesheet.Fills ?? new Fills();
        this.stylesheet.Borders = this.stylesheet.Borders ?? new Borders();
        this.stylesheet.CellFormats = this.stylesheet.CellFormats ?? new CellFormats();
    }

    /// <summary>
    /// Creates a custom style for a cell with specified font settings, text color, and optional alignment.
    /// </summary>
    /// <param name="fontName">
    /// The name of the font to be used in the style (e.g., "Arial").
    /// </param>
    /// <param name="fontSize">
    /// The size of the font in points. Must be greater than zero.
    /// </param>
    /// <param name="hexColor">
    /// The hexadecimal color code for the text (e.g., "000000" for black, "FF0000" for red).
    /// </param>
    /// <param name="horizontalAlignment">
    /// Optional. Specifies the horizontal alignment of the text within the cell. 
    /// Acceptable values are defined in the <see cref="HorizontalAlignment"/> enumeration.
    /// </param>
    /// <param name="verticalAlignment">
    /// Optional. Specifies the vertical alignment of the text within the cell. 
    /// Acceptable values are defined in the <see cref="VerticalAlignment"/> enumeration.
    /// </param>
    /// <returns>
    /// A <see cref="uint"/> representing the index of the created style in the workbook's stylesheet.
    /// </returns>
    /// <exception cref="ArgumentNullException">
    /// Thrown if <paramref name="fontName"/> or <paramref name="hexColor"/> is null or empty.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown if <paramref name="fontSize"/> is less than or equal to zero.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// Thrown if <paramref name="hexColor"/> is not a valid hexadecimal color code.
    /// </exception>
    /// <example>
    /// <code>
    /// StyleUtility styleUtility = new StyleUtility(stylesPart);
    /// uint styleIndex = styleUtility.CreateStyle("Calibri", 11, "FF0000", HorizontalAlignment.Center, VerticalAlignment.Center);
    /// Console.WriteLine($"Custom style index: {styleIndex}");
    /// </code>
    /// </example>
    /// <remarks>
    /// This method creates a custom cell format with the specified font, color, and alignment settings 
    /// and adds it to the workbook's stylesheet. If no alignment is specified, default alignment settings are used.
    /// </remarks>
    public uint CreateStyle(string fontName, double fontSize, string hexColor, HorizontalAlignment? horizontalAlignment = null,
    VerticalAlignment? verticalAlignment = null)
    {
        // Validate inputs
        if (string.IsNullOrEmpty(fontName))
            throw new ArgumentNullException(nameof(fontName));
        if (fontSize <= 0)
            throw new ArgumentOutOfRangeException(nameof(fontSize), "Font size must be greater than zero");
        if (string.IsNullOrEmpty(hexColor) || !IsHexColor(hexColor))
            throw new ArgumentException("Invalid hex color", nameof(hexColor));

        // Create and append Font
        var font = new Font(
            new FontSize() { Val = fontSize },
            new Color() { Rgb = new HexBinaryValue() { Value = hexColor } },
            new FontName() { Val = fontName }
        );
        this.stylesheet.Fonts.Append(font);

        // Create and append Fill
        var fill = new Fill(new PatternFill() { PatternType = PatternValues.None });
        this.stylesheet.Fills.Append(fill);

        // Create and append Border
        var border = new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder());
        this.stylesheet.Borders.Append(border);

        // Create and append CellFormat
        var cellFormat = new CellFormat
        {
            FontId = new UInt32Value((uint)this.stylesheet.Fonts.ChildElements.Count - 1),
            FillId = new UInt32Value((uint)this.stylesheet.Fills.ChildElements.Count - 1),
            BorderId = new UInt32Value((uint)this.stylesheet.Borders.ChildElements.Count - 1)
        };

        // Add alignment if specified
        if (horizontalAlignment.HasValue || verticalAlignment.HasValue)
        {
            cellFormat.Alignment = new Alignment
            {
                Horizontal = horizontalAlignment.HasValue
                    ? ConvertToOpenXmlHorizontalAlignment(horizontalAlignment.Value)
                    : null,
                Vertical = verticalAlignment.HasValue
                    ? ConvertToOpenXmlVerticalAlignment(verticalAlignment.Value)
                    : null
            };
        }

        this.stylesheet.CellFormats.Append(cellFormat);

        return (uint)this.stylesheet.CellFormats.ChildElements.Count - 1;
    }

    /// <summary>
    /// Converts a <see cref="HorizontalAlignment"/> value to the corresponding Open XML horizontal alignment value.
    /// </summary>
    /// <param name="alignment">The <see cref="HorizontalAlignment"/> to convert.</param>
    /// <returns>
    /// The equivalent <see cref="HorizontalAlignmentValues"/> used by Open XML.
    /// </returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown if the specified <paramref name="alignment"/> is not a valid value.
    /// </exception>
    private HorizontalAlignmentValues ConvertToOpenXmlHorizontalAlignment(HorizontalAlignment alignment)
    {
        return alignment switch
        {
            HorizontalAlignment.Left => HorizontalAlignmentValues.Left,
            HorizontalAlignment.Center => HorizontalAlignmentValues.Center,
            HorizontalAlignment.Right => HorizontalAlignmentValues.Right,
            HorizontalAlignment.Justify => HorizontalAlignmentValues.Justify,
            HorizontalAlignment.Distributed => HorizontalAlignmentValues.Distributed,
            _ => throw new ArgumentOutOfRangeException(nameof(alignment), "Invalid horizontal alignment")
        };
    }

    /// <summary>
    /// Converts a <see cref="VerticalAlignment"/> value to the corresponding Open XML vertical alignment value.
    /// </summary>
    /// <param name="alignment">The <see cref="VerticalAlignment"/> to convert.</param>
    /// <returns>
    /// The equivalent <see cref="VerticalAlignmentValues"/> used by Open XML.
    /// </returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown if the specified <paramref name="alignment"/> is not a valid value.
    /// </exception>
    private VerticalAlignmentValues ConvertToOpenXmlVerticalAlignment(VerticalAlignment alignment)
    {
        return alignment switch
        {
            VerticalAlignment.Top => VerticalAlignmentValues.Top,
            VerticalAlignment.Center => VerticalAlignmentValues.Center,
            VerticalAlignment.Bottom => VerticalAlignmentValues.Bottom,
            VerticalAlignment.Justify => VerticalAlignmentValues.Justify,
            VerticalAlignment.Distributed => VerticalAlignmentValues.Distributed,
            _ => throw new ArgumentOutOfRangeException(nameof(alignment), "Invalid vertical alignment")
        };
    }

    private static bool IsHexColor(string color)
    {
        return System.Text.RegularExpressions.Regex.IsMatch(color, "^(#)?([0-9a-fA-F]{3})([0-9a-fA-F]{3})?$");
    }


    /// <summary>
    /// Saves the current state of the workbook's stylesheet.
    /// </summary>
    /// <remarks>
    /// This method ensures that any changes made to the workbook's styles are persisted 
    /// by saving the updated stylesheet to the <see cref="WorkbookStylesPart"/>.
    /// </remarks>
    /// <example>
    /// <code>
    /// StyleUtility styleUtility = new StyleUtility(stylesPart);
    /// styleUtility.SaveStylesheet();
    /// Console.WriteLine("Stylesheet saved successfully.");
    /// </code>
    /// </example>
    public void SaveStylesheet()
    {
        if (this.stylesPart.Stylesheet == null)
        {
            this.stylesPart.Stylesheet = this.stylesheet;
        }
        this.stylesheet.Save();
    }
}
