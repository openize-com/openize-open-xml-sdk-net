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
    public class DoubleArrowFacade
    {
        
        private long _x;
        private long _y;
        private long _width;
        private long _height;
        private P.Shape _DoubleArrowShape;
        private SlidePart _AssociatedSlidePart;// Store the P.Shape as a private field
        private int _ShapeIndex;
        private AnimationType _Animation = AnimationType.None;

        private String _BackgroundColor;
        private ListFacade _TextList = null;
       
        public long X { get => _x; set => _x = value; }
        public long Y { get => _y; set => _y = value; }
        public long Width { get => _width; set => _width = value; }
        public long Height { get => _height; set => _height = value; }
        public P.Shape DoubleArrowShape { get => _DoubleArrowShape; set => _DoubleArrowShape = value; }
        public SlidePart AssociatedSlidePart { get => _AssociatedSlidePart; set => _AssociatedSlidePart = value; }
        public int ShapeIndex { get => _ShapeIndex; set => _ShapeIndex = value; }
        public string BackgroundColor { get => _BackgroundColor; set => _BackgroundColor = value; }
        public ListFacade TextList { get => _TextList; set => _TextList = value; }
        public AnimationType Animation { get => _Animation; set => _Animation = value; }

        public DoubleArrowFacade()
        {
            
           

        }

        public DoubleArrowFacade WithBackgroundColor(String backgroundColor)
        {
            BackgroundColor = backgroundColor;
            return this;
        }
       

        public DoubleArrowFacade WithPosition(long x, long y)
        {
            X = x;
            Y = y;
            return this;
        }

        public DoubleArrowFacade WithSize(long width, long height)
        {
            Width = width;
            Height = height;
            return this;
        }

        public P.Shape CreateShape()
        {
            
            P.Shape shape1 = new P.Shape();
            shape1.Append(CreateNonVisualShapeProperties());
            if (_BackgroundColor is null)
                shape1.Append(CreateShapeProperties(X, Y, Width, Height));
            else
                shape1.Append(CreateShapeProperties(X, Y, Width, Height, BackgroundColor));
            shape1.Append(CreateShapeStyle());
            shape1.Append(CreateTextBody());


            return shape1;
        }
      
       

        private P.ShapeStyle CreateShapeStyle()
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
        private P.ShapeProperties CreateShapeProperties(long x, long y, long width, long height, string rgbColorHex = "Transparent")
        {
            P.ShapeProperties shapeProperties1 = new P.ShapeProperties();

            D.Transform2D transform2D1 = new D.Transform2D();
            D.Offset offset1 = new D.Offset() { X = x, Y = y };
            D.Extents extents1 = new D.Extents() { Cx = width, Cy = height };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            D.PresetGeometry presetGeometry1 = new D.PresetGeometry() { Preset = D.ShapeTypeValues.StraightConnector1 };
            D.AdjustValueList adjustValueList1 = new D.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);


            D.SolidFill solidFill1 = new D.SolidFill();
            if (rgbColorHex != "Transparent")
            {
                D.RgbColorModelHex rgbColorModelHex1 = new D.RgbColorModelHex() { Val = rgbColorHex };
                solidFill1.Append(rgbColorModelHex1);
            }
            D.Outline outline1 = new D.Outline() ;
            D.HeadEnd headEnd1 = new D.HeadEnd() { Type = D.LineEndValues.Triangle };
            D.TailEnd tailEnd1 = new D.TailEnd() { Type = D.LineEndValues.Triangle };

            outline1.Append(headEnd1);
            outline1.Append(tailEnd1);


            D.SolidFill solidFill2 = new D.SolidFill();
            /* D.SchemeColor schemeColor1 = new D.SchemeColor() { Val = D.SchemeColorValues.Text1 };

             solidFill2.Append(schemeColor1);*/

           // outline1.Append(new NoFill());

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            if (rgbColorHex != "Transparent")
                shapeProperties1.Append(solidFill1);
            else
                shapeProperties1.Append(new NoFill());
            shapeProperties1.Append(outline1);

            return shapeProperties1;

        }
        private P.TextBody CreateTextBody()
        {
            P.TextBody textBody1 = new P.TextBody();
            D.BodyProperties bodyProperties1 = new D.BodyProperties() { RightToLeftColumns = false, Anchor = D.TextAnchoringTypeValues.Center };
            D.ListStyle listStyle1 = new D.ListStyle();

            D.Paragraph paragraph1 = new D.Paragraph();
            D.ParagraphProperties paragraphProperties1 = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Center };
            D.EndParagraphRunProperties endParagraphRunProperties1 = new D.EndParagraphRunProperties() { Language = "en-US" };

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);           

            return textBody1;
        }
        

        private P.NonVisualShapeProperties CreateNonVisualShapeProperties()
        {
            P.NonVisualShapeProperties nonVisualShapeProperties1 = new P.NonVisualShapeProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties1 = new P.NonVisualDrawingProperties() { Id = (UInt32Value)9U, Name = "Line 8" };
            P.NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new P.NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties1);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);
            nonVisualShapeProperties1.Append(applicationNonVisualDrawingProperties1);

            return nonVisualShapeProperties1;
        }

        public void UpdateShape()
        {
            if (DoubleArrowShape == null)
            {
                throw new InvalidOperationException("Shape has not been created yet. Call CreateShape() first.");
            }
            var alignmentType = TextAlignmentTypeValues.Justified;



            // Update the properties of the existing shape
            DoubleArrowShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id = (UInt32Value)5U;
            DoubleArrowShape.NonVisualShapeProperties.NonVisualDrawingProperties.Name = "Text Box 1";
            DoubleArrowShape.NonVisualShapeProperties.NonVisualShapeDrawingProperties = new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true });
            DoubleArrowShape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties(new PlaceholderShape());
            if (Width != 0)
            {
                DoubleArrowShape.ShapeProperties.Transform2D = new D.Transform2D()
                {
                    Offset = new D.Offset() { X = X, Y = Y },
                    Extents = new D.Extents() { Cx = Width, Cy = Height }
                };
            }
            if (!String.IsNullOrEmpty(_BackgroundColor))
            {
                if (_BackgroundColor == "Transparent")
                {
                    if (DoubleArrowShape.ShapeProperties.Descendants<NoFill>().FirstOrDefault() == null)
                    {
                        if (DoubleArrowShape.ShapeProperties.Descendants<SolidFill>().FirstOrDefault() != null)
                            DoubleArrowShape.ShapeProperties.Descendants<SolidFill>().FirstOrDefault().Remove();
                    }
                    else
                    {
                        DoubleArrowShape.ShapeProperties.Append(new NoFill());
                    }

                }
                else
                {
                    if (DoubleArrowShape.ShapeProperties.Descendants<NoFill>().FirstOrDefault() != null)
                    {
                        DoubleArrowShape.ShapeProperties.Descendants<NoFill>().FirstOrDefault().Remove();
                    }
                    var fill = DoubleArrowShape.ShapeProperties.Descendants<SolidFill>().FirstOrDefault();

                    if (fill != null)
                    {
                        fill.RemoveAllChildren();
                        fill.Append(new RgbColorModelHex() { Val = _BackgroundColor });
                    }
                    else
                    {
                        DoubleArrowShape.ShapeProperties.Append(new SolidFill(new RgbColorModelHex() { Val = _BackgroundColor }));
                    }

                }
            }

            var existingParagraphText = DoubleArrowShape.TextBody.Descendants<Run>().FirstOrDefault();
            DoubleArrowShape.TextBody.Elements<Paragraph>().FirstOrDefault().RemoveAllChildren();
            if (alignmentType != TextAlignmentTypeValues.Justified)
                DoubleArrowShape.TextBody.Elements<Paragraph>().FirstOrDefault().Append(new ParagraphProperties() { Alignment = alignmentType });
            DoubleArrowShape.TextBody.Elements<Paragraph>().FirstOrDefault().Append(existingParagraphText);

            var runProperties = DoubleArrowShape.TextBody.Descendants<RunProperties>().FirstOrDefault();

           
            var latinFont = runProperties.Elements<LatinFont>().FirstOrDefault();

          

            var solidFill = runProperties.Elements<SolidFill>().FirstOrDefault();

           }


        public void RemoveShape(SlidePart slidePart)
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
                var shapesToRemove = commonSlideData.ShapeTree.Elements<P.Shape>().Where(shape => shape == DoubleArrowShape).ToList();

                foreach (var shape in shapesToRemove)
                {
                    shape.Remove();
                }
            }
        }
        public void RemoveShape(P.Shape shape)
        {
            shape.Remove();
        }

        private static bool IsDoubleArrowShape(P.Shape shape)
        {
            var shapeProperties = shape.ShapeProperties;
            if (shapeProperties != null)
            {
                var presetGeometry = shapeProperties.GetFirstChild<D.PresetGeometry>();
                if (presetGeometry != null && presetGeometry.Preset == D.ShapeTypeValues.StraightConnector1)
                {
                    return true;
                }
            }
            return false;
        }
        // Method to populate List<rectangleShapeFacade> from a collection of P.Shape
        public static List<DoubleArrowFacade> PopulateDoubleArrowShapes(SlidePart slidePart)
        {
            IEnumerable<P.Shape> shapes = slidePart.Slide.CommonSlideData.ShapeTree.Elements<P.Shape>();
            var rectangleShapes = new List<DoubleArrowFacade>();
            var shapeIndex = 0;
            foreach (var shape in shapes)
            {
                if (IsDoubleArrowShape(shape))
                {
                    var rectangleShapeFacade = new DoubleArrowFacade
                    {
                        DoubleArrowShape = shape, // Store the P.Shape in the private field



                        X = GetXFromShape(shape),
                        Y = GetYFromShape(shape),
                        Width = GetWidthFromShape(shape),
                        Height = GetHeightFromShape(shape),
                        AssociatedSlidePart = slidePart,
                        ShapeIndex = shapeIndex
                    };

                    rectangleShapes.Add(rectangleShapeFacade);
                    shapeIndex += 1;
                }
            }

            return rectangleShapes;
        }

        private static string GetTextFromrectangleShape(P.Shape rectangleShape)
        {
            if (rectangleShape.TextBody != null)
            {
                return rectangleShape.TextBody.Descendants<Text>().FirstOrDefault()?.Text;
            }
            return null;
        }

      
          private static string GetFontFamilyFromrectangleShape(P.Shape rectangleShape)
        {
            var paragraph = rectangleShape.TextBody?.Elements<Paragraph>().FirstOrDefault();

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
        private static string GetColorFromrectangleShape(P.Shape rectangleShape)
        {
            var paragraph = rectangleShape.TextBody?.Elements<Paragraph>().FirstOrDefault();

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

        private static TextAlignment GetAlignmentFromDoubleArrowShape(P.Shape rectangleShape)
        {
            var alignment = rectangleShape.TextBody?.Descendants<Paragraph>().FirstOrDefault();
            if (alignment != null)
            {
                alignment = null;
            }
            var paragraphProperties = rectangleShape?.Descendants<P.TextBody>()?.FirstOrDefault()?.Descendants<Paragraph>()?
                   .FirstOrDefault();
            TextAlignmentTypeValues alignmentType = rectangleShape.TextBody.Descendants<ParagraphProperties>().FirstOrDefault()?.Alignment ?? TextAlignmentTypeValues.Justified;
            return ConvertAlignmentFromTypeValues(alignmentType);
        }

        private static long GetXFromShape(P.Shape shape)
        {
            return shape.ShapeProperties?.Transform2D?.Offset?.X ?? 0;
        }

        private static long GetYFromShape(P.Shape shape)
        {
            return shape.ShapeProperties?.Transform2D?.Offset?.Y ?? 0;
        }

        private static long GetWidthFromShape(P.Shape shape)
        {
            return shape.ShapeProperties?.Transform2D?.Extents?.Cx ?? 0;
        }

        private static long GetHeightFromShape(P.Shape shape)
        {
            return shape.ShapeProperties?.Transform2D?.Extents?.Cy ?? 0;
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
