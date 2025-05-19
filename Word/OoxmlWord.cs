using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DF = DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using FF = Openize.Words.IElements;
using OWD = OpenXML.Words.Data;
using OT = OpenXML.Templates;
using Openize.Words;
using System.Xml.Linq;

namespace OpenXML.Words
{
    internal class OwDocument
    {
        private WordprocessingDocument _pkgDocument;
        private WP.Body _wpBody;
        private MemoryStream _ms;
        private MainDocumentPart _mainPart;
        private List<int> _IDs;
        private NumberingDefinitionsPart _numberingPart;
        private readonly object _lockObject = new object();
        private OwDocument()
        {
            lock (_lockObject)
            {
                try
                {
                    _ms = new MemoryStream();
                    _pkgDocument = WordprocessingDocument.Create(_ms, DF.WordprocessingDocumentType.Document, true);
                    _mainPart = _pkgDocument.AddMainDocumentPart();
                    _mainPart.Document = new WP.Document();
                    var tmp = new OT.DefaultTemplate();
                    tmp.CreateMainDocumentPart(_mainPart);

                    _numberingPart = _mainPart.NumberingDefinitionsPart;

                    if (_numberingPart != null)
                    {
                        _IDs = new List<int>();
                        foreach (var abstractNum in _numberingPart.Numbering.Elements<WP.AbstractNum>())
                        {
                            _IDs.Add(abstractNum.AbstractNumberId);
                        }
                    }
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Initialize OOXML Element(s)");
                    throw new Openize.Words.OpenizeException(errorMessage, ex);
                }
            }
        }
        private OwDocument(WordprocessingDocument pkg)
        {
            lock (_lockObject)
            {
                try
                {
                    _pkgDocument = pkg;
                    _mainPart = pkg.MainDocumentPart;
                    _numberingPart = _mainPart.NumberingDefinitionsPart;

                    if (_numberingPart != null)
                    {
                        _IDs = new List<int>();
                        foreach (var abstractNum in _numberingPart.Numbering.Elements<WP.AbstractNum>())
                        {
                            _IDs.Add(abstractNum.AbstractNumberId);
                        }
                    }
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Initialize OOXML Element(s)");
                    throw new Openize.Words.OpenizeException(errorMessage, ex);
                }
            }
        }

        internal List<int> IDs
        {
            get { return new List<int>(_IDs); }
        }

        internal NumberingDefinitionsPart NumberingPart
        {
            get { return _numberingPart; }
        }

        internal MainDocumentPart MainPart
        {
            get { return _mainPart; }
        }

        #region Create Core Properties for OpenXML Word Document
        internal void CreateProperties(WordprocessingDocument pkgDocument,DocumentProperties documentProperties)
        {
            var corePart = pkgDocument.CoreFilePropertiesPart;
            
            if (corePart != null)
            {
                pkgDocument.DeletePart(corePart);
            }
            var dictCoreProp = new Dictionary<string, string>
            {
                ["Title"] = documentProperties.Title,
                ["Subject"] = documentProperties.Subject,
                ["Creator"] = documentProperties.Creator,
                ["Keywords"] = documentProperties.Keywords,
                ["Description"] = documentProperties.Description,
                ["LastModifiedBy"] = documentProperties.LastModifiedBy,
                ["Revision"] = documentProperties.Revision,
                ["Created"] = documentProperties.Created,
                ["Modified"] = documentProperties.Modified
            };
            var coreProperties = new OT.CoreProperties();
            coreProperties.CreateCoreFilePropertiesPart(pkgDocument.AddCoreFilePropertiesPart(), dictCoreProp);

            var customPart = pkgDocument.CustomFilePropertiesPart;
            if (customPart != null)
            {
                pkgDocument.DeletePart(customPart);
            }
            var customProperties = new OT.CustomProperties();
            customProperties.CreateExtendedFilePropertiesPart(pkgDocument.AddExtendedFilePropertiesPart());
        }
        #endregion

        public static OwDocument CreateInstance()
        {
            return new OwDocument();
        }

        public static OwDocument CreateInstance(WordprocessingDocument pkg)
        {
            return new OwDocument(pkg);
        }

        internal void CreateDocument(List<FF.IElement> lst,DocumentProperties documentProperties)
        {
            try
            {
                _wpBody = _mainPart.Document.Body;

                if (_wpBody == null)
                    throw new Openize.Words.OpenizeException("Package or Document or Body is null", new NullReferenceException());

                var sectionProperties = _wpBody.Elements<WP.SectionProperties>().FirstOrDefault();

                foreach (var element in lst)
                {
                    switch (element)
                    {
                        case FF.Paragraph ffP:
                            {
                                var para = OoxmlParagraph.CreateInstance(_IDs,_numberingPart).CreateParagraph(ffP);
                                _wpBody.InsertBefore(para, sectionProperties);
                                break;
                            }
                        case FF.Table ffTable:
                            {
                                var table = OoxmlTable.CreateInstance(_IDs, _numberingPart).CreateTable(ffTable);
                                _wpBody.InsertBefore(table, sectionProperties);
                                break;
                            }
                        case FF.Image ffImg:
                            {
                                var para = OoxmlImage.CreateInstance(_mainPart).CreateImage(ffImg);
                                _wpBody.InsertBefore(para, sectionProperties);
                                break;
                            }

                        case FF.Shape ffShape:
                            {
                                var para = OoxmlShape.CreateInstance().CreateShape(ffShape);
                                _wpBody.InsertBefore(para, sectionProperties);
                                break;
                            }

                        case FF.GroupShape ffGroupShape:
                            {
                                var para = OoxmlGroupShape.CreateInstance().
                                    CreateGroupShape(ffGroupShape);
                                _wpBody.InsertBefore(para, sectionProperties);
                                break;
                            }

                    }
                }
                CreateProperties(_pkgDocument,documentProperties);
            }
            catch (Exception ex)
            {
                var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Initialize OOXML Element(s)");
                throw new Openize.Words.OpenizeException(errorMessage, ex);
            }

        }
        internal List<FF.IElement> LoadDocument(Stream stream)
        {
            lock (_lockObject)
            {
                try
                {
                    _pkgDocument = WordprocessingDocument.Open(stream, true);
                    if (_pkgDocument.MainDocumentPart?.Document?.Body == null)
                        throw new Openize.Words.OpenizeException("Package or Document or Body is null",
                            new NullReferenceException());
                    OWD.OoxmlDocData.CreateInstance(_pkgDocument);

                    _mainPart = _pkgDocument.MainDocumentPart;
                    _numberingPart = _mainPart.NumberingDefinitionsPart;
                    _wpBody = _pkgDocument.MainDocumentPart.Document.Body;

                    //LoadProperties(_pkgDocument);

                    var sequence = 1;
                    var elements = new List<FF.IElement>();

                    foreach (var element in _wpBody.Elements())
                    {
                        switch (element)
                        {
                            case WP.Paragraph wpPara:
                                {
                                    var drawingFound = false;

                                    foreach (var drawing in wpPara.Descendants<WP.Drawing>())
                                    {
                                        var image = OoxmlImage.CreateInstance(_mainPart).LoadImage(drawing, sequence);
                                        if (image != null)
                                        {
                                            elements.Add(image);
                                            sequence++;
                                            drawingFound = true;
                                        }
                                        else
                                        {
                                            var inline = drawing.Inline;

                                            if (inline != null)
                                            {
                                                // Extract shape information from inline
                                                var graphic = inline.Graphic;
                                                var graphicData = graphic.GraphicData;

                                                if (graphicData.Uri.Value == "http://schemas.microsoft.com/office/word/2010/wordprocessingShape")
                                                {
                                                    var shape = OoxmlShape.CreateInstance().LoadShape(
                                                        graphicData, sequence);
                                                    if (shape != null)
                                                    {
                                                        elements.Add(shape);
                                                        sequence++;
                                                        drawingFound = true;
                                                    }
                                                }
                                                else if (graphicData.Uri.Value == "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup")
                                                {
                                                    var groupShape = OoxmlGroupShape.CreateInstance().
                                                        LoadGroupShape(graphicData, sequence);
                                                    if (groupShape != null)
                                                    {
                                                        elements.Add(groupShape);
                                                        sequence++;
                                                        drawingFound = true;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    if (!drawingFound)
                                    {
                                        elements.Add(OoxmlParagraph.CreateInstance(_IDs, _numberingPart).LoadParagraph(wpPara, sequence));
                                        sequence++;
                                    }
                                    break;
                                }

                            case WP.Drawing drawing:
                                {

                                    var image = OoxmlImage.CreateInstance(_mainPart).LoadImage(drawing, sequence);
                                    if (image != null)
                                    {
                                        elements.Add(OoxmlImage.CreateInstance(_mainPart).LoadImage(drawing, sequence));
                                        sequence++;
                                    }
                                    else
                                    {
                                        elements.Add(new FF.Unknown { ElementId = sequence });
                                        sequence++;
                                    }

                                    break;
                                }
                            case WP.Table wpTable:
                                elements.Add(OoxmlTable.CreateInstance(_IDs, _numberingPart).LoadTable(wpTable, sequence));
                                sequence++;
                                break;
                            case WP.SectionProperties wpSection:
                                elements.Add(LoadSection(wpSection, sequence));
                                sequence++;
                                break;
                            default:
                                elements.Add(new FF.Unknown { ElementId = sequence });
                                sequence++;
                                break;
                        }
                    }

                    return elements;

                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Load OOXML Elements");
                    throw new Openize.Words.OpenizeException(errorMessage, ex);
                }
            }
        }

        internal DocumentProperties LoadProperties()
        {
            var corePart = _pkgDocument.CoreFilePropertiesPart;
            DocumentProperties documentProperties = new DocumentProperties();

            if (corePart != null)
            {
                // Load the XML document from the CoreFilePropertiesPart
                XDocument coreXml = XDocument.Load(corePart.GetStream());

                // Define the namespaces used in core properties
                XNamespace dc = "http://purl.org/dc/elements/1.1/";
                XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
                XNamespace dcterms = "http://purl.org/dc/terms/";
                // Extract metadata values using the appropriate XML elements
                documentProperties.Title = coreXml.Descendants(dc + "title").FirstOrDefault()?.Value;
                documentProperties.Subject = coreXml.Descendants(dc + "subject").FirstOrDefault()?.Value;
                documentProperties.Description = coreXml.Descendants(dc + "description").FirstOrDefault()?.Value;
                documentProperties.Creator = coreXml.Descendants(dc + "creator").FirstOrDefault()?.Value;
                documentProperties.Keywords = coreXml.Descendants(cp + "keywords").FirstOrDefault()?.Value;
                documentProperties.LastModifiedBy = coreXml.Descendants(cp + "lastModifiedBy").FirstOrDefault()?.Value;
                documentProperties.Revision = coreXml.Descendants(cp + "revision").FirstOrDefault()?.Value;
                documentProperties.Created = coreXml.Descendants(dcterms + "created").FirstOrDefault()?.Value;
                documentProperties.Modified = coreXml.Descendants(dcterms + "modified").FirstOrDefault()?.Value;
            }
            else
            {
                //Console.WriteLine("Core properties not found.");
            }
            return documentProperties;
        }
        internal FF.Section LoadSection(WP.SectionProperties sectPr, int id)
        {
            lock (_lockObject)
            {
                try
                {
                    var section = new FF.Section
                    {
                        ElementId = id
                    };
                    if (sectPr != null)
                    {
                        var pageSize = sectPr.Elements<WP.PageSize>().FirstOrDefault();
                        if (pageSize != null)
                        {
                            section.PageSize = new FF.PageSize
                            {
                                Height = int.Parse(pageSize.Height),
                                Width = int.Parse(pageSize.Width),
                                Orientation = pageSize.Orient,
                            };
                        }

                        var pageMargin = sectPr.Elements<WP.PageMargin>().FirstOrDefault();
                        if (pageMargin != null)
                        {
                            section.PageMargin = new FF.PageMargin
                            {
                                Top = int.Parse(pageMargin.Top),
                                Right = int.Parse(pageMargin.Right),
                                Bottom = int.Parse(pageMargin.Bottom),
                                Left = int.Parse(pageMargin.Left),
                                Header = int.Parse(pageMargin.Header),
                                Footer = int.Parse(pageMargin.Footer),
                            };
                        }
                    }

                    return section;
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Load Section");
                    throw new Openize.Words.OpenizeException(errorMessage, ex);
                }
            }
        }
        internal FF.ElementStyles LoadStyles()
        {
            lock (_lockObject)
            {
                try
                {
                    var elementStyles = new FF.ElementStyles();
                    var themePart = _mainPart.ThemePart;
                    if (themePart != null)
                    {
                        var theme = themePart.Theme;
                        foreach (var fontScheme in theme.Elements())
                        {
                            foreach (var latinFont in fontScheme.Descendants<A.LatinFont>())
                            {
                                elementStyles.ThemeFonts.Add(latinFont.Typeface);
                            }
                        }

                        foreach (var fontScheme in theme.Elements())
                        {
                            var fonts = fontScheme.Descendants<A.SupplementalFont>();

                            foreach (var font in fonts)
                            {
                                if (font.Typeface != null)
                                {
                                    elementStyles.ThemeFonts.Add(font.Typeface);
                                }
                            }
                        }
                    }

                    var fontTablePart = _mainPart.FontTablePart;
                    if (fontTablePart != null)
                    {
                        var fontTable = fontTablePart.Fonts.Elements<WP.Font>();

                        foreach (var font in fontTable)
                        {
                            elementStyles.TableFonts.Add(font.Name);
                        }
                    }

                    var styleDefinitionsPart = _mainPart.StyleDefinitionsPart;

                    if (styleDefinitionsPart == null) return elementStyles;
                    var styles = styleDefinitionsPart.Styles;
                    if (styles != null)
                    {
                        foreach (var style in styles.Elements<WP.Style>())
                        {
                            if (style.Type != null && style.Type == WP.StyleValues.Paragraph)
                            {
                                elementStyles.ParagraphStyles.Add(style.StyleId);
                            }

                            if (style.Type != null && style.Type == WP.StyleValues.Table)
                            {
                                elementStyles.TableStyles.Add(style.StyleId);
                            }
                        }
                    }

                    return elementStyles;
                }
                catch (Exception ex)
                {
                    return null;
                }
            }
        }

        #region Save OpenXML Word Document to Stream
        internal void SaveDocument(Stream stream)
        {
            lock (_lockObject)
            {
                try
                {
                    _pkgDocument.Clone(stream);
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Dispose of managed resources (if any)
                if (_pkgDocument != null)
                {
                    _pkgDocument.Dispose();
                    _pkgDocument = null;
                }
            }
            // Dispose of unmanaged resources
            if (_ms == null) return;
            _ms.Dispose();
            _ms = null;
        }
        #endregion
    }
}
