using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection;

namespace Openize.Cells
{
    /// <summary>
    /// Provides extension methods for managing worksheet display properties and settings.
    /// </summary>
    public static class WorksheetPropertiesExtensions
    {
        /// <summary>
        /// Sets the zoom level for the worksheet view.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="zoomPercentage">The zoom percentage (10-400).</param>
        /// <returns>The worksheet for method chaining.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when zoomPercentage is out of valid range (10-400).</exception>
        public static Worksheet SetZoom(this Worksheet worksheet, int zoomPercentage)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            if (zoomPercentage < 10 || zoomPercentage > 400)
                throw new ArgumentOutOfRangeException(nameof(zoomPercentage), "Zoom percentage must be between 10 and 400.");

            // Just use the worksheet's public API
            var part = GetWorksheetPart(worksheet);
            if (part != null)
            {
                // Only update the ZoomScale without touching other elements
                EnsureSheetViewElement(part);

                var sheetViews = part.Worksheet.GetFirstChild<SheetViews>();
                var sheetView = sheetViews.GetFirstChild<SheetView>();
                if (sheetView != null)
                {
                    sheetView.ZoomScale = (uint)zoomPercentage;
                }
            }

            return worksheet;
        }

        /// <summary>
        /// Sets the default column width for the worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="width">The default column width in characters.</param>
        /// <returns>The worksheet for method chaining.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when width is not positive.</exception>
        public static Worksheet SetDefaultColumnWidth(this Worksheet worksheet, double width)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            if (width <= 0)
                throw new ArgumentOutOfRangeException(nameof(width), "Default column width must be positive.");

            var part = GetWorksheetPart(worksheet);
            if (part != null)
            {
                EnsureSheetFormatPrElement(part);

                var sheetFormatPr = part.Worksheet.GetFirstChild<SheetFormatProperties>();
                if (sheetFormatPr != null)
                {
                    sheetFormatPr.DefaultColumnWidth = width;
                }
            }

            return worksheet;
        }

        /// <summary>
        /// Sets the default row height for the worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="height">The default row height in points.</param>
        /// <returns>The worksheet for method chaining.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when height is not positive.</exception>
        public static Worksheet SetDefaultRowHeight(this Worksheet worksheet, double height)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            if (height <= 0)
                throw new ArgumentOutOfRangeException(nameof(height), "Default row height must be positive.");

            var part = GetWorksheetPart(worksheet);
            if (part != null)
            {
                EnsureSheetFormatPrElement(part);

                var sheetFormatPr = part.Worksheet.GetFirstChild<SheetFormatProperties>();
                if (sheetFormatPr != null)
                {
                    sheetFormatPr.DefaultRowHeight = height;
                    sheetFormatPr.CustomHeight = true;
                }
            }

            return worksheet;
        }

        /// <summary>
        /// Sets whether to show or hide formulas in the worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="show">True to show formulas; false to hide formulas.</param>
        /// <returns>The worksheet for method chaining.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        public static Worksheet ShowFormulas(this Worksheet worksheet, bool show)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            var part = GetWorksheetPart(worksheet);
            if (part != null)
            {
                EnsureSheetViewElement(part);

                var sheetViews = part.Worksheet.GetFirstChild<SheetViews>();
                var sheetView = sheetViews.GetFirstChild<SheetView>();
                if (sheetView != null)
                {
                    sheetView.ShowFormulas = show;
                }
            }

            return worksheet;
        }

        /// <summary>
        /// Sets whether to show or hide gridlines in the worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="show">True to show gridlines; false to hide gridlines.</param>
        /// <returns>The worksheet for method chaining.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        public static Worksheet ShowGridlines(this Worksheet worksheet, bool show)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            var part = GetWorksheetPart(worksheet);
            if (part != null)
            {
                EnsureSheetViewElement(part);

                var sheetViews = part.Worksheet.GetFirstChild<SheetViews>();
                var sheetView = sheetViews.GetFirstChild<SheetView>();
                if (sheetView != null)
                {
                    sheetView.ShowGridLines = show;
                }
            }

            return worksheet;
        }

        /// <summary>
        /// Sets whether to show or hide row and column headers in the worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="show">True to show headers; false to hide headers.</param>
        /// <returns>The worksheet for method chaining.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        public static Worksheet ShowRowColumnHeaders(this Worksheet worksheet, bool show)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            var part = GetWorksheetPart(worksheet);
            if (part != null)
            {
                EnsureSheetViewElement(part);

                var sheetViews = part.Worksheet.GetFirstChild<SheetViews>();
                var sheetView = sheetViews.GetFirstChild<SheetView>();
                if (sheetView != null)
                {
                    sheetView.ShowRowColHeaders = show;
                }
            }

            return worksheet;
        }

        /// <summary>
        /// Sets whether to show or hide zero values in the worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="show">True to show zero values; false to hide zero values.</param>
        /// <returns>The worksheet for method chaining.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        public static Worksheet ShowZeroValues(this Worksheet worksheet, bool show)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            var part = GetWorksheetPart(worksheet);
            if (part != null)
            {
                EnsureSheetViewElement(part);

                var sheetViews = part.Worksheet.GetFirstChild<SheetViews>();
                var sheetView = sheetViews.GetFirstChild<SheetView>();
                if (sheetView != null)
                {
                    sheetView.ShowZeros = show;
                }
            }

            return worksheet;
        }

        /// <summary>
        /// Sets whether the worksheet is displayed right-to-left.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="rightToLeft">True for right-to-left display; false for left-to-right display.</param>
        /// <returns>The worksheet for method chaining.</returns>
        /// <exception cref="ArgumentNullException">Thrown when worksheet is null.</exception>
        public static Worksheet SetRightToLeft(this Worksheet worksheet, bool rightToLeft)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            var part = GetWorksheetPart(worksheet);
            if (part != null)
            {
                EnsureSheetViewElement(part);

                var sheetViews = part.Worksheet.GetFirstChild<SheetViews>();
                var sheetView = sheetViews.GetFirstChild<SheetView>();
                if (sheetView != null)
                {
                    sheetView.RightToLeft = rightToLeft;
                }
            }

            return worksheet;
        }

        /// <summary>
        /// Ensures that the SheetFormatProperties element exists in the worksheet
        /// </summary>
        private static void EnsureSheetFormatPrElement(WorksheetPart worksheetPart)
        {
            var worksheet = worksheetPart.Worksheet;
            var sheetFormatPr = worksheet.GetFirstChild<SheetFormatProperties>();

            if (sheetFormatPr == null)
            {
                sheetFormatPr = new SheetFormatProperties();

                // Find the correct position for SheetFormatProperties according to the schema
                var sheetViews = worksheet.GetFirstChild<SheetViews>();
                if (sheetViews != null)
                {
                    worksheet.InsertAfter(sheetFormatPr, sheetViews);
                }
                else
                {
                    var sheetPr = worksheet.GetFirstChild<SheetProperties>();
                    if (sheetPr != null)
                    {
                        worksheet.InsertAfter(sheetFormatPr, sheetPr);
                    }
                    else
                    {
                        worksheet.PrependChild(sheetFormatPr);
                    }
                }
            }
        }

        /// <summary>
        /// Ensures that the SheetViews and SheetView elements exist in the worksheet
        /// </summary>
        private static void EnsureSheetViewElement(WorksheetPart worksheetPart)
        {
            var worksheet = worksheetPart.Worksheet;
            var sheetViews = worksheet.GetFirstChild<SheetViews>();

            if (sheetViews == null)
            {
                sheetViews = new SheetViews();

                // Find the correct position for SheetViews according to the schema
                var sheetPr = worksheet.GetFirstChild<SheetProperties>();
                if (sheetPr != null)
                {
                    worksheet.InsertAfter(sheetViews, sheetPr);
                }
                else
                {
                    worksheet.PrependChild(sheetViews);
                }

                // Add a default SheetView
                var newSheetView = new SheetView { WorkbookViewId = 0 };
                sheetViews.Append(newSheetView);
            }
            else
            {
                var existingSheetView = sheetViews.GetFirstChild<SheetView>();
                if (existingSheetView == null)
                {
                    var newSheetView = new SheetView { WorkbookViewId = 0 };
                    sheetViews.Append(newSheetView);
                }
            }
        }

        /// <summary>
        /// Gets the OpenXML worksheet part from a Worksheet object.
        /// </summary>
        private static WorksheetPart GetWorksheetPart(Worksheet worksheet)
        {
            // Use reflection to access the _worksheetPart field from the Worksheet class
            var fieldInfo = worksheet.GetType().GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
            if (fieldInfo == null)
                return null;

            return fieldInfo.GetValue(worksheet) as WorksheetPart;
        }
    }
}