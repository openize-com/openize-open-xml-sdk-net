using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using Text = DocumentFormat.OpenXml.Drawing.Text;
using NonVisualGroupShapeProperties = DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using Openize.Slides.Common.Enumerations;
using Openize.Slides.Common;
using System.Collections.Generic;
using System;
using DocumentFormat.OpenXml.Drawing.Charts;
using SD = System.Data;
using System.Data;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;



namespace Openize.Slides.Facade
{
    public class TableFacade
    {
        private SlidePart _AssociatedSlidePart;
        private D.Table _Table;
        private SD.DataTable _SD_DTable = null;
        private string _RelationshipId;
        private int _TableIndex;
        private Int64Value _x;
        private Int64Value _y;
        private Int64Value _width;
        private Int64Value _height;
        private Stylings _CellStylings;
        private P.GraphicFrame _TableGraphicFrame;
        private String _TableStyle;

        public SlidePart TableSlidePart { get => _AssociatedSlidePart; set => _AssociatedSlidePart = value; }
        public D.Table Table { get => _Table; set => _Table = value; }
        public string RelationshipId { get => _RelationshipId; set => _RelationshipId = value; }
        public int TableIndex { get => _TableIndex; set => _TableIndex = value; }
        public Int64Value X { get => _x; set => _x = value; }
        public Int64Value Y { get => _y; set => _y = value; }
        public Int64Value Width { get => _width; set => _width = value; }
        public Int64Value Height { get => _height; set => _height = value; }
        public Stylings CellStylings { get => _CellStylings; set => _CellStylings = value; }

        public SD.DataTable SD_DTable { get => _SD_DTable; set => _SD_DTable = value; }
        public P.GraphicFrame TableGraphicFrame { get => _TableGraphicFrame; set => _TableGraphicFrame = value; }
        public string TableStyle { get => _TableStyle; set => _TableStyle = value; }

        public TableFacade()
        {

        }

        public void GenerateTable(SlidePart slidePart, SD.DataTable table)
        {
            slidePart.Slide.CommonSlideData.ShapeTree.Append(CreateTable(table));
        }
        private P.GraphicFrame CreateTable(SD.DataTable dataTable)
        {
            P.GraphicFrame graphicFrame = new P.GraphicFrame();

            // Non-visual properties for the graphic frame
            P.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties = new P.NonVisualGraphicFrameProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "table 8" };

            P.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties = new P.NonVisualGraphicFrameDrawingProperties();
            D.GraphicFrameLocks graphicFrameLocks = new D.GraphicFrameLocks() { NoGrouping = true };

            nonVisualGraphicFrameDrawingProperties.Append(graphicFrameLocks);

            P.ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();

            P.ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList = new P.ApplicationNonVisualDrawingPropertiesExtensionList();

            P.ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension = new P.ApplicationNonVisualDrawingPropertiesExtension() { Uri = "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}" };

            P14.ModificationId modificationId = new P14.ModificationId() { Val = (UInt32Value)3331517366U };
            modificationId.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            applicationNonVisualDrawingPropertiesExtension.Append(modificationId);

            applicationNonVisualDrawingPropertiesExtensionList.Append(applicationNonVisualDrawingPropertiesExtension);

            applicationNonVisualDrawingProperties.Append(applicationNonVisualDrawingPropertiesExtensionList);

            nonVisualGraphicFrameProperties.Append(nonVisualDrawingProperties);
            nonVisualGraphicFrameProperties.Append(nonVisualGraphicFrameDrawingProperties);
            nonVisualGraphicFrameProperties.Append(applicationNonVisualDrawingProperties);

            // Transform properties
            Transform transform = new Transform();
            D.Offset offset = new D.Offset() { X = _x, Y = _y };
            D.Extents extents = new D.Extents() { Cx = _width, Cy = _height };

            transform.Append(offset);
            transform.Append(extents);

            // Graphic
            D.Graphic graphic = new D.Graphic();
            D.GraphicData graphicData = new D.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" };

            // Table properties
            D.TableProperties tableProperties = new D.TableProperties() { FirstRow = true, BandRow = true };
            D.TableStyleId tableStyleId = PowerPointTableStyles.PowerPointTableStylesMap.GetTableStyleId(TableStyle);

            tableProperties.Append(tableStyleId);
            // Table
            D.Table table = CreateTableData(dataTable, tableProperties);

            graphicData.Append(table);
            graphic.Append(graphicData);

            graphicFrame.Append(nonVisualGraphicFrameProperties);
            graphicFrame.Append(transform);
            graphicFrame.Append(graphic);

            return graphicFrame;
        }

        public void UpdateTable(SD.DataTable dtable)
        {
            // Get the existing elements from the GraphicFrame
            _width = _TableGraphicFrame.Descendants<Transform>()
            .FirstOrDefault()?.Descendants<Extents>()
            .FirstOrDefault()?.Cx ?? 0;
            GraphicData existingGraphicData = _TableGraphicFrame.Descendants<GraphicData>().FirstOrDefault();
            var tableProperties = _TableGraphicFrame.Descendants<D.TableProperties>().FirstOrDefault();

            if(_TableStyle != null)
            {
                tableProperties.Descendants<TableStyleId>().FirstOrDefault().Remove();
                var tableStyleId = PowerPointTableStyles.PowerPointTableStylesMap.GetTableStyleId(TableStyle);
                tableProperties.Append(tableStyleId);
            }
            // Remove the existing elements from the GraphicFrame
            _TableGraphicFrame.Descendants<D.TableProperties>().FirstOrDefault().Remove();
            _TableGraphicFrame.Descendants<D.Table>().FirstOrDefault().Remove();

            // Create the new Table
            D.Table newTable = CreateTableData(dtable, tableProperties);
            existingGraphicData.Append(newTable);

        }

        private D.Table CreateTableData(SD.DataTable dtable, D.TableProperties tableProperties)
        {
            D.Table table = new D.Table();
            // Table grid
            D.TableGrid tableGrid = new D.TableGrid();
            var extId = 0;
            foreach (var column in dtable.Columns)
            {
                D.GridColumn gridColumn = new D.GridColumn() { Width = _width / dtable.Columns.Count };
                gridColumn.Append(CreateColumnExtension("2000" + extId));
                tableGrid.Append(gridColumn);
                extId += 1;
            }
            // Append elements to the table
            table.Append(tableProperties);
            table.Append(tableGrid);

            // Rows
            D.TableRow headerRow = CreateHeaderRow(dtable.Columns);
            table.Append(headerRow);
            foreach (SD.DataRow row in dtable.Rows)
            {
                D.TableRow dataRow = CreateDataRow(row);
                table.Append(dataRow);
            }
            return table;
        }

        private D.TableRow CreateHeaderRow(SD.DataColumnCollection columns)
        {
            D.TableRow tableRow = new D.TableRow() { Height = 370840L };

            foreach (SD.DataColumn column in columns)
            {
                D.TableCell tableCell;

                if (_CellStylings.Equals(default(Stylings)))
                    tableCell = CreateTableCell(column.ColumnName);
                else
                    tableCell = CreateTableCell(column.ColumnName, _CellStylings);

                tableRow.Append(tableCell);
            }

            // Add ExtensionList with OpenXmlUnknownElement for rowId
            tableRow.Append(CreateRowExtension("10000"));

            return tableRow;
        }
        private D.TableRow CreateDataRow(SD.DataRow row)
        {
            D.TableRow tableRow = new D.TableRow() { Height = 370840L };

            foreach (var item in row.ItemArray)
            {
                D.TableCell tableCell = CreateTableCell(item.ToString());
                tableRow.Append(tableCell);
            }

            // Add ExtensionList with OpenXmlUnknownElement for rowId
            tableRow.Append(CreateRowExtension("10001"));

            return tableRow;
        }
        private D.TableRow CreateDataRow(string id, string text)
        {
            // Create a data row with two cells
            D.TableRow tableRow = CreateTableRow(CreateTableCell(id), CreateTableCell(text));

            // Add ExtensionList with OpenXmlUnknownElement for rowId
            tableRow.Append(CreateRowExtension("10001"));

            return tableRow;
        }
        public D.TableCell CreateTableCell(string text_styling)
        {
            string text = "";
            var stylings = new Stylings();

            if (text_styling.Contains(';'))
            {
                text = text_styling.Split(';')[0];
                stylings = Utility.DeserializeStyling(text_styling.Split(';')[1]);
                if (stylings.FontSize > 0)
                {
                    return CreateTableCell(text, stylings);
                }
            }
            else
            {
                text = text_styling;
            }


            D.TableCell tableCell = new D.TableCell();



            D.TextBody textBody = new D.TextBody();
            D.BodyProperties bodyProperties = new D.BodyProperties();
            D.ListStyle listStyle = new D.ListStyle();

            D.Paragraph paragraph = new D.Paragraph();
            D.Run run = new D.Run();
            D.RunProperties runProperties = new D.RunProperties() { Language = "en-US", Dirty = false };
            D.Text cellText = new D.Text() { Text = text };

            run.Append(runProperties);
            run.Append(cellText);

            D.EndParagraphRunProperties endParagraphRunProperties = new D.EndParagraphRunProperties() { Language = "en-AS", Dirty = false };

            paragraph.Append(run);
            paragraph.Append(endParagraphRunProperties);

            textBody.Append(bodyProperties);
            textBody.Append(listStyle);
            textBody.Append(paragraph);

            D.TableCellProperties tableCellProperties = new D.TableCellProperties();

            tableCell.Append(textBody);
            tableCell.Append(tableCellProperties);

            return tableCell;
        }

        public D.TableCell CreateTableCell(string text, Stylings styling)
        {
            // Create a TableCell with the specified text
            D.TableCell tableCell = new D.TableCell();

            D.TextBody textBody = new D.TextBody();
            D.BodyProperties bodyProperties = new D.BodyProperties();
            D.ListStyle listStyle = new D.ListStyle();
            D.Paragraph paragraph = new D.Paragraph(
                                    new D.ParagraphProperties() { Alignment = ConvertAlignmentToTypeValues(styling.Alignment) },
                                    new D.Run(
                                        new D.RunProperties(new D.SolidFill(new D.RgbColorModelHex() { Val = styling.TextColor }),
                                        new D.LatinFont() { Typeface = styling.FontFamily })
                                        { FontSize = styling.FontSize * 100, Dirty = false },
                                        new D.Text() { Text = text }
                                    )
                                );
            //D.ParagraphProperties paragraphProperties1 = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Center };
            D.EndParagraphRunProperties endParagraphRunProperties1 = new D.EndParagraphRunProperties() { Language = "es-ES" };

            //paragraph.Append(paragraphProperties1);
            paragraph.Append(endParagraphRunProperties1);

            textBody.Append(bodyProperties);
            textBody.Append(listStyle);
            textBody.Append(paragraph);

            D.TableCellProperties tableCellProperties = new D.TableCellProperties();

            tableCell.Append(textBody);
            tableCell.Append(tableCellProperties);

            return tableCell;
        }
        private D.TableRow CreateTableRow(params D.TableCell[] cells)
        {
            // Create a TableRow with the specified cells
            D.TableRow tableRow = new D.TableRow() { Height = _height / 2 };

            foreach (D.TableCell cell in cells)
            {
                tableRow.Append(cell);
            }


            return tableRow;
        }

        private static D.ExtensionList CreateColumnExtension(string colIdValue)
        {
            // Create ExtensionList with OpenXmlUnknownElement for columnId
            D.ExtensionList extensionList = new D.ExtensionList();
            D.Extension extension = new D.Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" };
            
            OpenXmlUnknownElement unknownElement = new OpenXmlUnknownElement($"<a16:colId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"{colIdValue}\" />");
            extension.Append(unknownElement);
            extensionList.Append(extension);

            return extensionList;
        }

        private D.ExtensionList CreateRowExtension(string rowIdValue)
        {
            // Create ExtensionList with OpenXmlUnknownElement for rowId
            D.ExtensionList extensionList = new D.ExtensionList();
            D.Extension extension = new D.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };
            OpenXmlUnknownElement unknownElement = new OpenXmlUnknownElement($"<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"{rowIdValue}\" />");
            extension.Append(unknownElement);
            extensionList.Append(extension);

            return extensionList;
        }


        public static List<TableFacade> PopulateTables(SlidePart slidePart)
        {
            List<TableFacade> tableFacades = new List<TableFacade>();

            var slide = slidePart.Slide;
            var shapeTree = slide.Descendants<ShapeTree>().FirstOrDefault();
            if (shapeTree != null)
            {
                foreach (var frame in shapeTree.Descendants<P.GraphicFrame>())
                {
                    var table = frame.Descendants<D.Table>().FirstOrDefault();
                    if (table != null)
                    {
                        var tableFacade = new TableFacade();
                        tableFacade.SD_DTable = CreateTable(table);
                        tableFacade.TableGraphicFrame = frame;

                        tableFacades.Add(tableFacade);
                    }

                }

            }


            return tableFacades;
        }

        private static System.Data.DataTable CreateTable(D.Table dtable)
        {
            // Create a new DataTable
            System.Data.DataTable dataTable = new System.Data.DataTable();

            // Retrieve table properties
            D.TableProperties tableProperties = dtable.TableProperties;

            // Get table grid
            D.TableGrid tableGrid = dtable.TableGrid;

            // Extract column information from table grid
            List<string> columnNames = new List<string>();
            foreach (D.TableRow headerRow in dtable.Elements<D.TableRow>().Take(1))
            {
                foreach (D.TableCell dcell in headerRow.Elements<D.TableCell>())
                {
                    // Extract text from the cell and add it to column names
                    string cellText = ExtractCellText(dcell);
                    columnNames.Add(cellText);
                }
            }

            // Add columns to DataTable based on column names
            foreach (string columnName in columnNames)
            {
                dataTable.Columns.Add(columnName, typeof(string)); // Adjust column types if needed
            }

            // Iterate through rows in the D.Table, skipping the header row
            foreach (D.TableRow drow in dtable.Elements<D.TableRow>().Skip(1))
            {
                // Create a new DataRow for the DataTable
                DataRow dataRow = dataTable.NewRow();

                // Index to keep track of the current cell's position
                int columnIndex = 0;

                // Iterate through cells in the D.TableRow
                foreach (D.TableCell dcell in drow.Elements<D.TableCell>())
                {
                    // Extract text from the cell
                    string cellText = ExtractCellText(dcell);

                    // Set the cell value in the DataRow
                    dataRow[columnIndex++] = cellText;
                }

                // Add the DataRow to the DataTable
                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }


        // Helper method to extract text from a D.TableCell
        private static string ExtractCellText(D.TableCell cell)
        {
            var text1 = cell.Descendants<D.Text>().FirstOrDefault();
            return text1.Text;
        }

        private static TextAlignmentTypeValues ConvertAlignmentToTypeValues(TextAlignment alignment)
        {
            switch (alignment)
            {
                case TextAlignment.Left:
                    return TextAlignmentTypeValues.Left;
                case TextAlignment.Center:
                    return TextAlignmentTypeValues.Center;
                case TextAlignment.Right:
                    return TextAlignmentTypeValues.Right;
                case TextAlignment.None:
                    return TextAlignmentTypeValues.Justified;
                default:
                    throw new ArgumentOutOfRangeException(nameof(alignment), alignment, null);
            }
        }

        private static TextAlignment ConvertAlignmentFromTypeValues(TextAlignmentTypeValues alignmentType)
        {
            if (alignmentType.Equals(TextAlignmentTypeValues.Left))
                return TextAlignment.Left;
            if (alignmentType.Equals(TextAlignmentTypeValues.Center))
                return TextAlignment.Center;
            if (alignmentType.Equals(TextAlignmentTypeValues.Right))
                return TextAlignment.Right;
            if (alignmentType.Equals(TextAlignmentTypeValues.Justified))
                return TextAlignment.None;

            throw new ArgumentOutOfRangeException(nameof(alignmentType), alignmentType, null);
        }





    }
}
