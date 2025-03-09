using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Openize.Cells
{
    /// <summary>
    /// Provides extension methods for managing worksheet tab colors in Excel workbooks.
    /// </summary>
    public static class WorksheetColorExtensions
    {
        /// <summary>
        /// Sets the tab color for a worksheet using RGB values.
        /// </summary>
        /// <param name="worksheet">The worksheet to set the tab color for.</param>
        /// <param name="red">The red component (0-255).</param>
        /// <param name="green">The green component (0-255).</param>
        /// <param name="blue">The blue component (0-255).</param>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when color components are outside the valid range.</exception>
        public static void SetTabColor(this Worksheet worksheet, byte red, byte green, byte blue)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            string hexColor = $"{red:X2}{green:X2}{blue:X2}";
            SetTabColorByHex(worksheet, hexColor);
        }

        /// <summary>
        /// Sets the tab color for a worksheet using an HTML-style hex color code.
        /// </summary>
        /// <param name="worksheet">The worksheet to set the tab color for.</param>
        /// <param name="hexColor">The hex color code (e.g., "FF0000" for red).</param>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        /// <exception cref="ArgumentException">Thrown when hexColor is null, empty, or invalid.</exception>
        public static void SetTabColorByHex(this Worksheet worksheet, string hexColor)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            if (string.IsNullOrEmpty(hexColor))
                throw new ArgumentException("Hex color cannot be null or empty.", nameof(hexColor));

            // Strip # if present
            if (hexColor.StartsWith("#"))
                hexColor = hexColor.Substring(1);

            // Validate hex color format
            if (!IsValidHexColor(hexColor))
                throw new ArgumentException("Invalid hex color format. Expected format: RRGGBB", nameof(hexColor));

            // Get access to the internal OpenXML worksheet
            var openXmlWorksheetPart = GetOpenXmlWorksheetPart(worksheet);
            if (openXmlWorksheetPart == null)
                return;

            var openXmlWorksheet = openXmlWorksheetPart.Worksheet;

            // Get or create the SheetProperties element
            SheetProperties sheetProperties = openXmlWorksheet.GetFirstChild<SheetProperties>();
            if (sheetProperties == null)
            {
                sheetProperties = new SheetProperties();
                openXmlWorksheet.InsertAt(sheetProperties, 0);
            }

            // Get or create the TabColor element
            TabColor tabColor = sheetProperties.GetFirstChild<TabColor>();
            if (tabColor == null)
            {
                tabColor = new TabColor();
                sheetProperties.AppendChild(tabColor);
            }

            // Set the RGB value
            tabColor.Rgb = hexColor;

            // Save the changes
            openXmlWorksheet.Save();
        }

        /// <summary>
        /// Gets the tab color of a worksheet as RGB values.
        /// </summary>
        /// <param name="worksheet">The worksheet to get the tab color from.</param>
        /// <param name="red">Output parameter to receive the red component.</param>
        /// <param name="green">Output parameter to receive the green component.</param>
        /// <param name="blue">Output parameter to receive the blue component.</param>
        /// <returns>True if a tab color is set; otherwise, false.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        public static bool GetTabColor(this Worksheet worksheet, out byte red, out byte green, out byte blue)
        {
            red = 0;
            green = 0;
            blue = 0;

            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            string hexColor = GetTabColorAsHex(worksheet);
            if (string.IsNullOrEmpty(hexColor))
                return false;

            // Parse the hex color
            red = Convert.ToByte(hexColor.Substring(0, 2), 16);
            green = Convert.ToByte(hexColor.Substring(2, 2), 16);
            blue = Convert.ToByte(hexColor.Substring(4, 2), 16);

            return true;
        }

        /// <summary>
        /// Gets the tab color of a worksheet as a hex color code.
        /// </summary>
        /// <param name="worksheet">The worksheet to get the tab color from.</param>
        /// <returns>The tab color as a hex color code, or null if no tab color is set.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        public static string GetTabColorAsHex(this Worksheet worksheet)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            // Get access to the internal OpenXML worksheet
            var openXmlWorksheetPart = GetOpenXmlWorksheetPart(worksheet);
            if (openXmlWorksheetPart == null)
                return null;

            var openXmlWorksheet = openXmlWorksheetPart.Worksheet;

            // Check if tab color is set
            SheetProperties sheetProperties = openXmlWorksheet.GetFirstChild<SheetProperties>();
            if (sheetProperties == null)
                return null;

            TabColor tabColor = sheetProperties.GetFirstChild<TabColor>();
            if (tabColor == null || string.IsNullOrEmpty(tabColor.Rgb))
                return null;

            return tabColor.Rgb;
        }

        /// <summary>
        /// Removes the tab color from a worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet to remove the tab color from.</param>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        public static void RemoveTabColor(this Worksheet worksheet)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            // Get access to the internal OpenXML worksheet
            var openXmlWorksheetPart = GetOpenXmlWorksheetPart(worksheet);
            if (openXmlWorksheetPart == null)
                return;

            var openXmlWorksheet = openXmlWorksheetPart.Worksheet;

            // Check if tab color is set
            SheetProperties sheetProperties = openXmlWorksheet.GetFirstChild<SheetProperties>();
            if (sheetProperties == null)
                return;

            TabColor tabColor = sheetProperties.GetFirstChild<TabColor>();
            if (tabColor != null)
            {
                tabColor.Remove();
                openXmlWorksheet.Save();
            }
        }

        /// <summary>
        /// Validates a hex color code.
        /// </summary>
        /// <param name="hexColor">The hex color code to validate.</param>
        /// <returns>True if the hex color code is valid; otherwise, false.</returns>
        private static bool IsValidHexColor(string hexColor)
        {
            if (string.IsNullOrEmpty(hexColor))
                return false;

            // Standard hex color format is either RRGGBB or AARRGGBB
            return System.Text.RegularExpressions.Regex.IsMatch(hexColor, "^[0-9A-Fa-f]{6}([0-9A-Fa-f]{2})?$");
        }

        /// <summary>
        /// Gets the OpenXML worksheet part from a Worksheet object.
        /// </summary>
        /// <param name="worksheet">The worksheet to get the OpenXML part from.</param>
        /// <returns>The OpenXML worksheet part, or null if not available.</returns>
        private static DocumentFormat.OpenXml.Packaging.WorksheetPart GetOpenXmlWorksheetPart(Worksheet worksheet)
        {
            // Use reflection to access the _worksheetPart field from the Worksheet class
            var fieldInfo = worksheet.GetType().GetField("_worksheetPart", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            if (fieldInfo == null)
                return null;

            return fieldInfo.GetValue(worksheet) as DocumentFormat.OpenXml.Packaging.WorksheetPart;
        }
    }
}