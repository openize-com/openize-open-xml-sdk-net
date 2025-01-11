using System;
using System.Collections.Generic;
using System.Linq;
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

namespace Openize.Slides.Facade
{
    public class TextShapeFacade
    {
        private string _Text;
        private int _fontSize;
        private TextAlignment _alignment;
        private long _x;
        private long _y;
        private long _width;
        private long _height;
        private P.Shape _textBoxShape;
        private SlidePart _AssociatedSlidePart;// Store the P.Shape as a private field
        private int _ShapeIndex;
        private String _FontFamily = "Calibri";
        private String _TextColor = "#000000";
        private List<TextSegmentFacade> _TextSegmentFacades;
        private String _BackgroundColor;
        private ListFacade _TextList = null;
        private AnimationType _Animation = AnimationType.None;
        public string Text { get => _Text; set => _Text = value; }
        // Public property to set and get font size in points
        public int FontSize
        {
            get
            {
                // Convert fontSizeTwentieths to points
                return (int)_fontSize / 20 / 2;
            }
            set
            {
                // Convert points to twentieths of a point and store in fontSizeTwentieths
                _fontSize = (int)(value * 20 * 2);
            }
        }
        public TextAlignment Alignment { get => _alignment; set => _alignment = value; }
        public long X { get => _x; set => _x = value; }
        public long Y { get => _y; set => _y = value; }
        public long Width { get => _width; set => _width = value; }
        public long Height { get => _height; set => _height = value; }
        public P.Shape TextBoxShape { get => _textBoxShape; set => _textBoxShape = value; }
        protected SlidePart AssociatedSlidePart { get => _AssociatedSlidePart; set => _AssociatedSlidePart = value; }
        public int ShapeIndex { get => _ShapeIndex; set => _ShapeIndex = value; }
        public string FontFamily { get => _FontFamily; set => _FontFamily = value; }
        public string TextColor { get => _TextColor; set => _TextColor = value; }
        public List<TextSegmentFacade> TextSegmentFacades { get => _TextSegmentFacades; set => _TextSegmentFacades = value; }
        public string BackgroundColor { get => _BackgroundColor; set => _BackgroundColor = value; }
        public ListFacade TextList { get => _TextList; set => _TextList = value; }
        public AnimationType Animation { get => _Animation; set => _Animation = value; }

        public TextShapeFacade ()
        {
            // Set default values or initialization logic if needed
            FontSize = 12; // Example default font size

        }

        public TextShapeFacade WithText (string text)
        {
            Text = text;
            return this;
        }

        public TextShapeFacade WithFontSize (int fontSize)
        {
            FontSize = fontSize;
            return this;
        }
        public TextShapeFacade WithFontFamily (String fontfamily)
        {
            FontFamily = fontfamily;
            return this;
        }
        public TextShapeFacade WithTextColor (String textColor)
        {
            TextColor = textColor;
            return this;
        }
        public TextShapeFacade WithBackgroundColor (String backgroundColor)
        {
            BackgroundColor = backgroundColor;
            return this;
        }
        public TextShapeFacade WithAlignment (TextAlignment alignment)
        {
            Alignment = alignment;
            return this;
        }

        public TextShapeFacade WithPosition (long x, long y)
        {
            X = x;
            Y = y;
            return this;
        }

        public TextShapeFacade WithSize (long width, long height)
        {
            Width = width;
            Height = height;
            return this;
        }

        public P.Shape CreateShape ()
        {
            TextAlignmentTypeValues alignmentType = ConvertAlignmentToTypeValues(Alignment);
            P.Shape shape1 = new P.Shape();
            shape1.Append(CreateNonVisualShapeProperties());
            if (_BackgroundColor is null)
                shape1.Append(CreateShapeProperties(X, Y, Width, Height));
            else
                shape1.Append(CreateShapeProperties(X, Y, Width, Height, BackgroundColor));
            shape1.Append(CreateShapeStyle());
            shape1.Append(CreateTextBody(_TextColor, _FontFamily, _fontSize, _Text, alignmentType));


            return shape1;
        }
        public P.Shape CreateListShape (List<String> listItems, ListFacade list)
        {
            TextAlignmentTypeValues alignmentType = ConvertAlignmentToTypeValues(Alignment);
            P.Shape shape1 = new P.Shape();
            shape1.Append(CreateNonVisualShapeProperties());
            if (_BackgroundColor is null)
                shape1.Append(CreateShapeProperties(X, Y, Width, Height));
            else
                shape1.Append(CreateShapeProperties(X, Y, Width, Height, BackgroundColor));
            shape1.Append(CreateShapeStyle());
            shape1.Append(list.CreateList(listItems, _TextColor, _FontFamily, _fontSize, new P.TextBody()));
            return shape1;
        }
        public P.Shape CreateShape (List<TextSegmentFacade> textSegmentFacades)
        {
            TextAlignmentTypeValues alignmentType = ConvertAlignmentToTypeValues(Alignment);
            P.Shape shape1 = new P.Shape();
            shape1.Append(CreateNonVisualShapeProperties());
            if (_BackgroundColor is null)
                shape1.Append(CreateShapeProperties(X, Y, Width, Height));
            else
                shape1.Append(CreateShapeProperties(X, Y, Width, Height, BackgroundColor));
            shape1.Append(CreateShapeStyle());
            shape1.Append(CreateTextBody(textSegmentFacades, alignmentType));


            return shape1;



        }

        private P.ShapeStyle CreateShapeStyle ()
        {
            P.ShapeStyle shapeStyle1 = new P.ShapeStyle();

            D.LineReference lineReference1 = new D.LineReference() { Index = (UInt32Value)2U };

            D.SchemeColor schemeColor2 = new D.SchemeColor() { Val = D.SchemeColorValues.Accent1 };
            D.Shade shade1 = new D.Shade() { Val = 50000 };

            schemeColor2.Append(shade1);

            lineReference1.Append(schemeColor2);

            D.FillReference fillReference1 = new D.FillReference() { Index = (UInt32Value)1U };
            D.SchemeColor schemeColor3 = new D.SchemeColor() { Val = D.SchemeColorValues.Accent1 };

            fillReference1.Append(schemeColor3);

            D.EffectReference effectReference1 = new D.EffectReference() { Index = (UInt32Value)0U };
            D.SchemeColor schemeColor4 = new D.SchemeColor() { Val = D.SchemeColorValues.Accent1 };

            effectReference1.Append(schemeColor4);

            D.FontReference fontReference1 = new D.FontReference() { Index = D.FontCollectionIndexValues.Minor };
            D.SchemeColor schemeColor5 = new D.SchemeColor() { Val = D.SchemeColorValues.Light1 };

            fontReference1.Append(schemeColor5);

            shapeStyle1.Append(lineReference1);
            shapeStyle1.Append(fillReference1);
            shapeStyle1.Append(effectReference1);
            shapeStyle1.Append(fontReference1);

            return shapeStyle1;
        }
        private P.ShapeProperties CreateShapeProperties (long x, long y, long width, long height, string rgbColorHex = "Transparent")
        {
            P.ShapeProperties shapeProperties1 = new P.ShapeProperties();

            D.Transform2D transform2D1 = new D.Transform2D();
            D.Offset offset1 = new D.Offset() { X = x, Y = y };
            D.Extents extents1 = new D.Extents() { Cx = width, Cy = height };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            D.PresetGeometry presetGeometry1 = new D.PresetGeometry() { Preset = D.ShapeTypeValues.Rectangle };
            D.AdjustValueList adjustValueList1 = new D.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);


            D.SolidFill solidFill1 = new D.SolidFill();
            if (rgbColorHex != "Transparent")
            {
                D.RgbColorModelHex rgbColorModelHex1 = new D.RgbColorModelHex() { Val = rgbColorHex };
                solidFill1.Append(rgbColorModelHex1);
            }
            D.Outline outline1 = new D.Outline() { Width = 12700 };

            D.SolidFill solidFill2 = new D.SolidFill();
            /* D.SchemeColor schemeColor1 = new D.SchemeColor() { Val = D.SchemeColorValues.Text1 };

             solidFill2.Append(schemeColor1);*/

            outline1.Append(new NoFill());

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            if (rgbColorHex != "Transparent")
                shapeProperties1.Append(solidFill1);
            else
                shapeProperties1.Append(new NoFill());
            shapeProperties1.Append(outline1);

            return shapeProperties1;

        }
        private P.TextBody CreateTextBody (string textColor, string font, int fontSize, string Text, TextAlignmentTypeValues alignmentType)
        {
            P.TextBody textBody1 = new P.TextBody();
            D.BodyProperties bodyProperties1 = new D.BodyProperties() { RightToLeftColumns = false, Anchor = D.TextAnchoringTypeValues.Center };
            D.ListStyle listStyle1 = new D.ListStyle();

            D.Paragraph paragraph1 = new D.Paragraph(
                         new ParagraphProperties() { Alignment = alignmentType },
                         new Run(
                             new RunProperties(new SolidFill(new RgbColorModelHex() { Val = _TextColor }),
                             new LatinFont() { Typeface = _FontFamily })
                             { FontSize = _fontSize, Dirty = false },
                             new Text() { Text = Text }
                         )
                     );
            D.ParagraphProperties paragraphProperties1 = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Center };
            D.EndParagraphRunProperties endParagraphRunProperties1 = new D.EndParagraphRunProperties() { Language = "es-ES" };

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);

            return textBody1;
        }
        private P.TextBody CreateTextBody (List<TextSegmentFacade> textSegmentFacades, TextAlignmentTypeValues alignmentType)
        {
            P.TextBody textBody1 = new P.TextBody();
            D.BodyProperties bodyProperties1 = new D.BodyProperties() { RightToLeftColumns = false, Anchor = D.TextAnchoringTypeValues.Center };
            D.ListStyle listStyle1 = new D.ListStyle();

            D.Paragraph paragraph1 = new D.Paragraph();
            D.ParagraphProperties paragraphProperties1 = new D.ParagraphProperties() { Alignment = alignmentType };
            // Append each run to the paragraph
            foreach (TextSegmentFacade textSegmentFacade in textSegmentFacades)
            {
                paragraph1.Append(textSegmentFacade);
            }

            D.EndParagraphRunProperties endParagraphRunProperties1 = new D.EndParagraphRunProperties() { Language = "es-ES" };

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);

            return textBody1;
        }

        private P.NonVisualShapeProperties CreateNonVisualShapeProperties ()
        {
            P.NonVisualShapeProperties nonVisualShapeProperties1 = new P.NonVisualShapeProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties1 = new P.NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "1 Shape Name" };
            P.NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new P.NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties1);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);
            nonVisualShapeProperties1.Append(applicationNonVisualDrawingProperties1);

            return nonVisualShapeProperties1;
        }

        public void UpdateShape ()
        {
            if (TextBoxShape == null)
            {
                throw new InvalidOperationException("Shape has not been created yet. Call CreateShape() first.");
            }
            var alignmentType = TextAlignmentTypeValues.Justified;

            if (Alignment != TextAlignment.None)
                alignmentType = ConvertAlignmentToTypeValues(Alignment);

            // Update the properties of the existing shape
            TextBoxShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id = (UInt32Value)5U;
            TextBoxShape.NonVisualShapeProperties.NonVisualDrawingProperties.Name = "Text Box 1";
            TextBoxShape.NonVisualShapeProperties.NonVisualShapeDrawingProperties = new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true });
            TextBoxShape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties(new PlaceholderShape());
            if (Width != 0)
            {
                TextBoxShape.ShapeProperties.Transform2D = new D.Transform2D()
                {
                    Offset = new D.Offset() { X = X, Y = Y },
                    Extents = new D.Extents() { Cx = Width, Cy = Height }
                };
            }
            if (!String.IsNullOrEmpty(_BackgroundColor))
            {
                if (_BackgroundColor == "Transparent")
                {
                    if (TextBoxShape.ShapeProperties.Descendants<NoFill>().FirstOrDefault() == null)
                    {
                        if (TextBoxShape.ShapeProperties.Descendants<SolidFill>().FirstOrDefault() != null)
                            TextBoxShape.ShapeProperties.Descendants<SolidFill>().FirstOrDefault().Remove();
                    }
                    else
                    {
                        TextBoxShape.ShapeProperties.Append(new NoFill());
                    }

                }
                else
                {
                    if (TextBoxShape.ShapeProperties.Descendants<NoFill>().FirstOrDefault() != null)
                    {
                        TextBoxShape.ShapeProperties.Descendants<NoFill>().FirstOrDefault().Remove();
                    }
                    var fill = TextBoxShape.ShapeProperties.Descendants<SolidFill>().FirstOrDefault();

                    if (fill != null)
                    {
                        fill.RemoveAllChildren();
                        fill.Append(new RgbColorModelHex() { Val = _BackgroundColor });
                    }
                    else
                    {
                        TextBoxShape.ShapeProperties.Append(new SolidFill(new RgbColorModelHex() { Val = _BackgroundColor }));
                    }

                }
            }

            var existingParagraphText = TextBoxShape.TextBody.Descendants<Run>().FirstOrDefault();
            TextBoxShape.TextBody.Elements<Paragraph>().FirstOrDefault().RemoveAllChildren();
            if (alignmentType != TextAlignmentTypeValues.Justified)
                TextBoxShape.TextBody.Elements<Paragraph>().FirstOrDefault().Append(new ParagraphProperties() { Alignment = alignmentType });
            TextBoxShape.TextBody.Elements<Paragraph>().FirstOrDefault().Append(existingParagraphText);

            var runProperties = TextBoxShape.TextBody.Descendants<RunProperties>().FirstOrDefault();

            if (_fontSize != 0)
            {
                runProperties.FontSize = _fontSize;
            }
            else if (runProperties.FontSize != null)
            {

            }
            else
            {
                // Extract font size from associated slide layout
                var slideLayout = _AssociatedSlidePart.SlideLayoutPart.SlideLayout;
                var defaultRunProperties = slideLayout.Descendants<DefaultRunProperties>().ToList()[_ShapeIndex];

                if (defaultRunProperties != null)
                {
                    runProperties.FontSize = defaultRunProperties.FontSize;
                }
            }
            var latinFont = runProperties.Elements<LatinFont>().FirstOrDefault();

            if (!string.IsNullOrEmpty(_FontFamily))
            {
                latinFont = new LatinFont() { Typeface = _FontFamily };
            }
            else if (latinFont != null)
            {

            }
            else
            {
                // Extract font family from associated slide layout
                var slideLayout = _AssociatedSlidePart.SlideLayoutPart.SlideLayout;
                var defaultRunProperties = slideLayout.Descendants<DefaultRunProperties>().ToList()[_ShapeIndex];
                var defaultLatinFont = defaultRunProperties.Elements<LatinFont>().FirstOrDefault();

                if (defaultRunProperties != null && defaultLatinFont != null)
                {
                    latinFont = new LatinFont() { Typeface = defaultLatinFont.Typeface };
                }
            }

            var solidFill = runProperties.Elements<SolidFill>().FirstOrDefault();

            if (!string.IsNullOrEmpty(_TextColor))
            {
                // Clear existing font-related properties
                runProperties.Elements<LatinFont>().ToList().ForEach(l => l.Remove());

                // Set text color using hex code (e.g., #FF0000 for red)
                solidFill = new SolidFill(new RgbColorModelHex() { Val = _TextColor });
                runProperties.AppendChild(solidFill);
            }
            else if (solidFill != null)
            {

            }
            else
            {
                // Extract color from associated slide layout
                var slideLayout = _AssociatedSlidePart.SlideLayoutPart.SlideLayout;
                var defaultRunProperties = slideLayout.Descendants<DefaultRunProperties>().ToList()[_ShapeIndex];

                if (defaultRunProperties != null)
                {
                    // Clear existing font-related properties
                    runProperties.Elements<LatinFont>().ToList().ForEach(l => l.Remove());

                    // Set text color using hex code (e.g., #FF0000 for red)
                    solidFill = new SolidFill(new RgbColorModelHex() { Val = _TextColor });
                    runProperties.AppendChild(solidFill);
                }
            }


            TextBoxShape.TextBody.Elements<Paragraph>().FirstOrDefault().Elements<Run>().FirstOrDefault().Text.Text = Text;
        }


        public void RemoveShape (SlidePart slidePart)
        {
            // Ensure slidePart is not null
            if (slidePart == null)
            {
                throw new ArgumentNullException("slidePart cannot be null.");
            }

            // Find the ShapeTree in CommonSlideData
            CommonSlideData commonSlideData = slidePart.Slide.CommonSlideData;
            if (commonSlideData != null && commonSlideData.ShapeTree != null)
            {
                // Remove the specified shape from the ShapeTree
                var shapesToRemove = commonSlideData.ShapeTree.Elements<P.Shape>().Where(shape => shape == TextBoxShape).ToList();

                foreach (var shape in shapesToRemove)
                {
                    shape.Remove();
                }
            }
        }
        public void RemoveShape (P.Shape shape)
        {
            shape.Remove();
        }

        // Method to populate List<TextShapeFacade> from a collection of P.Shape
        public static List<TextShapeFacade> PopulateTextShapes (SlidePart slidePart)
        {
            IEnumerable<P.Shape> shapes = slidePart.Slide.CommonSlideData.ShapeTree.Elements<P.Shape>();
            var textShapes = new List<TextShapeFacade>();
            var shapeIndex = 0;
            foreach (var shape in shapes)
            {

                var textShapeFacade = new TextShapeFacade
                {
                    TextBoxShape = shape, // Store the P.Shape in the private field
                    Text = GetTextFromTextShape(shape),
                    FontSize = GetFontSizeFromTextShape(shape),
                    TextColor = GetColorFromTextShape(shape),
                    Alignment = GetAlignmentFromTextShape(shape),
                    FontFamily = GetFontFamilyFromTextShape(shape),
                    X = GetXFromShape(shape),
                    Y = GetYFromShape(shape),
                    Width = GetWidthFromShape(shape),
                    Height = GetHeightFromShape(shape),
                    TextList = GetStyledList(shape),
                    AssociatedSlidePart = slidePart,
                    ShapeIndex = shapeIndex
                };

                textShapes.Add(textShapeFacade);
                shapeIndex += 1;
            }

            return textShapes;
        }

        private static string GetTextFromTextShape (P.Shape textShape)
        {
            if (textShape.TextBody != null)
            {
                return textShape.TextBody.Descendants<Text>().FirstOrDefault()?.Text;
            }
            return null;
        }

        private static int GetFontSizeFromTextShape (P.Shape textShape)
        {
            var paragraph = textShape.TextBody?.Elements<Paragraph>().FirstOrDefault();

            if (paragraph != null)
            {
                var defaultRunProperties = paragraph.Elements<ParagraphProperties>().FirstOrDefault()?.Elements<DefaultRunProperties>().FirstOrDefault();

                if (defaultRunProperties != null)
                {
                    return defaultRunProperties.FontSize;
                }
            }

            return 0;
        }
        private static ListFacade GetStyledList (P.Shape textShape)
        {
            if (textShape.TextBody.Descendants<BulletFont>().Count() > 0)
            {
                var bulletFont = textShape.TextBody.Descendants<BulletFont>().FirstOrDefault();
                var paragraphs = textShape.TextBody.Descendants<Paragraph>();
                var runProperties = paragraphs.FirstOrDefault().Descendants<RunProperties>().FirstOrDefault();
                var listFacade = new ListFacade();
                var listItems = new List<String>();
                listFacade.TextShape = textShape;
                listFacade.ListType = paragraphs.FirstOrDefault().Descendants<CharacterBullet>().Any() ? ListType.Bulleted : ListType.Numbered;
                listFacade.FontSize = runProperties.FontSize;
                listFacade.FontFamily = runProperties.Descendants<LatinFont>().FirstOrDefault().Typeface;
                listFacade.TextColor = runProperties.Descendants<RgbColorModelHex>().FirstOrDefault().Val;
                foreach (var paragraph in paragraphs)
                {

                    listItems.Add(paragraph.InnerText);
                }

                listFacade.ListItems = listItems;
                return listFacade;

            }
            return null;
        }
        private static string GetFontFamilyFromTextShape (P.Shape textShape)
        {
            var paragraph = textShape.TextBody?.Elements<Paragraph>().FirstOrDefault();

            if (paragraph != null)
            {
                var defaultRunProperties = paragraph.Elements<ParagraphProperties>().FirstOrDefault()?.Elements<DefaultRunProperties>().FirstOrDefault();

                if (defaultRunProperties != null)
                {
                    var latinFont = defaultRunProperties.Elements<LatinFont>().FirstOrDefault();

                    if (latinFont != null)
                    {
                        return latinFont.Typeface;
                    }
                }
            }

            return null; // or an appropriate default value for FontFamily
        }
        private static string GetColorFromTextShape (P.Shape textShape)
        {
            var paragraph = textShape.TextBody?.Elements<Paragraph>().FirstOrDefault();

            if (paragraph != null)
            {
                var defaultRunProperties = paragraph.Elements<ParagraphProperties>().FirstOrDefault()?.Elements<DefaultRunProperties>().FirstOrDefault();

                if (defaultRunProperties != null)
                {
                    var solidFill = defaultRunProperties.Elements<SolidFill>().FirstOrDefault();

                    if (solidFill != null)
                    {
                        var rgbColor = solidFill.Elements<RgbColorModelHex>().FirstOrDefault();

                        if (rgbColor != null)
                        {
                            return rgbColor.Val;
                        }
                    }
                }
            }

            return null; // or an appropriate default value for color code
        }

        private static TextAlignment GetAlignmentFromTextShape (P.Shape textShape)
        {
            var alignment = textShape.TextBody?.Descendants<Paragraph>().FirstOrDefault();
            if (alignment != null)
            {
                alignment = null;
            }
            var paragraphProperties = textShape?.Descendants<P.TextBody>()?.FirstOrDefault()?.Descendants<Paragraph>()?
                   .FirstOrDefault();
            TextAlignmentTypeValues alignmentType = textShape.TextBody.Descendants<ParagraphProperties>().FirstOrDefault()?.Alignment ?? TextAlignmentTypeValues.Justified;
            return ConvertAlignmentFromTypeValues(alignmentType);
        }

        private static long GetXFromShape (P.Shape shape)
        {
            return shape.ShapeProperties?.Transform2D?.Offset?.X ?? 0;
        }

        private static long GetYFromShape (P.Shape shape)
        {
            return shape.ShapeProperties?.Transform2D?.Offset?.Y ?? 0;
        }

        private static long GetWidthFromShape (P.Shape shape)
        {
            return shape.ShapeProperties?.Transform2D?.Extents?.Cx ?? 0;
        }

        private static long GetHeightFromShape (P.Shape shape)
        {
            return shape.ShapeProperties?.Transform2D?.Extents?.Cy ?? 0;
        }

        private static TextAlignmentTypeValues ConvertAlignmentToTypeValues (TextAlignment alignment)
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
