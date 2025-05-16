using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DF = DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DWG = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using DWS = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using FF = Openize.Words.IElements;
using OWD = OpenXML.Words.Data;
using OT = OpenXML.Templates;
using Openize.Words;
using System.Xml.Linq;

namespace OpenXML.Words
{
    internal class OoxmlTable
    {
        private readonly object _lockObject = new object();
        private List<int> _IDs;
        private NumberingDefinitionsPart _numberingPart;

        private OoxmlTable(List<int> IDs, NumberingDefinitionsPart numberingPart)
        {
            _IDs = IDs;
            _numberingPart = numberingPart;
        }

        internal static OoxmlTable CreateInstance(List<int> IDs, NumberingDefinitionsPart numberingPart)
        {
            return new OoxmlTable(IDs, numberingPart);
        }

        internal WP.Table CreateTable(FF.Table ffTable)
        {
            lock (_lockObject)
            {
                try
                {
                    var rows = ffTable.Rows.Count;
                    var cols = ffTable.Rows[0].Cells.Count;

                    var wpTable = new WP.Table(
                        new WP.TableProperties(
                            new WP.TableStyle() { Val = ffTable.Style } // Specify the TableStyle ID you want to apply
                        )
                    );
                    var tableGrid = new WP.TableGrid();
                    for (var i = 0; i < cols; i++)
                    {
                        if (ffTable.Column.Width > 0)
                            tableGrid.Append(new WP.GridColumn { Width = ffTable.Column.Width.ToString() });
                        else
                            tableGrid.Append(new WP.GridColumn());
                    }

                    wpTable.Append(tableGrid);

                    for (var i = 0; i < rows; i++)
                    {
                        var wpRow = new WP.TableRow();

                        for (var j = 0; j < cols; j++)
                        {
                            var wpCell = new WP.TableCell();
                            var ffCell = ffTable.Rows[i].Cells[j];
                            foreach (var ffPara in ffCell.Paragraphs)
                            {
                                //CreateParagraph(ffPara));
                                wpCell.Append(OoxmlParagraph.CreateInstance(_IDs, _numberingPart).CreateParagraph(ffPara));
                            }

                            wpRow.Append(wpCell);
                        }

                        wpTable.Append(wpRow);
                    }

                    return wpTable;
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Create Table");
                    throw new Openize.Words.OpenizeException(errorMessage, ex);
                }
            }
        }
        internal FF.Table LoadTable(WP.Table wpTable, int id)
        {
            lock (_lockObject)
            {
                try
                {
                    var ffRows = new List<FF.Row>();
                    foreach (var wpRow in wpTable.Elements<WP.TableRow>())
                    {
                        var ffRow = new FF.Row
                        {
                            Cells = new List<FF.Cell>()
                        };
                        foreach (var wpCell in wpRow.Elements<WP.TableCell>())
                        {
                            var ffParas = new List<FF.Paragraph>();
                            foreach (var paragraph in wpCell.Elements<WP.Paragraph>())
                            {
                                ffParas.Add(OoxmlParagraph.CreateInstance(_IDs, _numberingPart).LoadParagraph(paragraph, 0));
                            }

                            var ffCell = new FF.Cell { Paragraphs = ffParas };
                            ffRow.Cells.Add(ffCell);
                        }

                        ffRows.Add(ffRow);
                    }

                    var ffTable = new FF.Table
                    {
                        Rows = ffRows,
                        ElementId = id
                    };
                    var tableGrid = wpTable.Elements<WP.TableGrid>().FirstOrDefault();
                    if (tableGrid != null)
                    {
                        var gridColumn = tableGrid.Elements<WP.GridColumn>().FirstOrDefault();
                        ffTable.Column.Width = Convert.ToInt32(gridColumn.Width);
                    }

                    var tableProperties = wpTable.Descendants<WP.TableProperties>().FirstOrDefault();
                    if (tableProperties == null) return ffTable;
                    var tableStyle = tableProperties.TableStyle;
                    if (tableStyle != null)
                    {
                        ffTable.Style = tableStyle.Val;
                    }

                    return ffTable;
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Load Table");
                    throw new Openize.Words.OpenizeException(errorMessage, ex);
                }
            }
        }
    }
}
