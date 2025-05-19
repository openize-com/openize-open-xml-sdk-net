using System;
using DF = DocumentFormat.OpenXml;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DWS = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using FF = Openize.Words.IElements;
using OWD = OpenXML.Words.Data;

namespace OpenXML.Words
{
    internal class OoxmlShape
    {
        private readonly object _lockObject = new object();

        private OoxmlShape()
        {
        }

        internal static OoxmlShape CreateInstance()
        {
            return new OoxmlShape();
        }

        internal WP.Paragraph CreateShape(FF.Shape shape)
        {
            lock (_lockObject)
            {
                try
                {
                    var paragraph = new WP.Paragraph();
                    var run = new WP.Run();

                    var runProperties = new WP.RunProperties();
                    var noProof = new WP.NoProof();

                    runProperties.Append(noProof);

                    var alternateContent = new DF.AlternateContent();
                    alternateContent.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

                    var alternateContentChoice = new DF.AlternateContentChoice() { Requires = "wps" };
                    alternateContentChoice.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

                    var drawing = new WP.Drawing();
                    drawing.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                    var inline = new DW.Inline()
                    { DistanceFromTop = (DF.UInt32Value)0U, DistanceFromBottom = (DF.UInt32Value)0U, DistanceFromLeft = (DF.UInt32Value)0U, DistanceFromRight = (DF.UInt32Value)0U, AnchorId = "27EE2959", EditId = "551435BE" };
                    inline.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
                    inline.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");

                    var extent = new DW.Extent() { Cx = shape.Width * 9525, Cy = shape.Height * 9525 };
                    extent.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

                    var effectExtent = new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 13970L, BottomEdge = 13970L };
                    effectExtent.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

                    var docProperties = new DW.DocProperties() { Id = (DF.UInt32Value)1609145151U, Name = "Oval 1" };
                    docProperties.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

                    var nonVisualGraphicFrameDrawingProperties = new DW.NonVisualGraphicFrameDrawingProperties();
                    nonVisualGraphicFrameDrawingProperties.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

                    var graphic = new A.Graphic();
                    graphic.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                    var graphicData = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

                    var wordprocessingShape = new DWS.WordprocessingShape();
                    wordprocessingShape.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
                    var nonVisualDrawingShapeProperties = new DWS.NonVisualDrawingShapeProperties();

                    var shapeProperties = new DWS.ShapeProperties();

                    var transform2D = new A.Transform2D();
                    var offset = new A.Offset() { X = shape.X * 9525, Y = shape.Y * 9525 };
                    var extents = new A.Extents() { Cx = shape.Width * 9525, Cy = shape.Height * 9525 };

                    transform2D.Append(offset);
                    transform2D.Append(extents);

                    var presetGeometry = new A.PresetGeometry() { Preset = CreateShapeType(shape.Type) }; //A.ShapeTypeValues.Ellipse };
                    var adjustValueList = new A.AdjustValueList();

                    presetGeometry.Append(adjustValueList);
                    var outline = new A.Outline();

                    shapeProperties.Append(transform2D);
                    shapeProperties.Append(presetGeometry);
                    switch (shape.FillType)
                    {
                        case FF.ShapeFillType.Solid:
                            A.SolidFill solidFill = new A.SolidFill();
                            A.RgbColorModelHex rgbColor = new A.RgbColorModelHex()
                            { Val = shape.FillColors.Color1 };
                            solidFill.Append(rgbColor);
                            shapeProperties.Append(solidFill);
                            break;
                        case FF.ShapeFillType.Gradient:
                            shapeProperties.Append(CreateGradientFill(
                                shape.FillColors.Color1, shape.FillColors.Color2));
                            break;
                        case FF.ShapeFillType.Pattern:
                            shapeProperties.Append(CreatePatternFill(
                                shape.FillColors.Color1, shape.FillColors.Color2));
                            break;
                    }
                    shapeProperties.Append(outline);

                    var shapeStyle = new DWS.ShapeStyle();

                    var lineReference = new A.LineReference() { Index = (DF.UInt32Value)2U };

                    var schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
                    var shade = new A.Shade() { Val = 50000 };

                    schemeColor.Append(shade);

                    lineReference.Append(schemeColor);

                    var fillReference = new A.FillReference() { Index = (DF.UInt32Value)1U };
                    schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

                    fillReference.Append(schemeColor);

                    var effectReference = new A.EffectReference() { Index = (DF.UInt32Value)0U };
                    var rgbColorModelPercentage = new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

                    effectReference.Append(rgbColorModelPercentage);

                    var fontReference = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
                    schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

                    fontReference.Append(schemeColor);

                    shapeStyle.Append(lineReference);
                    shapeStyle.Append(fillReference);
                    shapeStyle.Append(effectReference);
                    shapeStyle.Append(fontReference);
                    var textBodyProperties = new DWS.TextBodyProperties() { Anchor = A.TextAnchoringTypeValues.Center };

                    wordprocessingShape.Append(nonVisualDrawingShapeProperties);
                    wordprocessingShape.Append(shapeProperties);
                    wordprocessingShape.Append(shapeStyle);
                    wordprocessingShape.Append(textBodyProperties);

                    graphicData.Append(wordprocessingShape);

                    graphic.Append(graphicData);

                    inline.Append(extent);
                    inline.Append(effectExtent);
                    inline.Append(docProperties);
                    inline.Append(nonVisualGraphicFrameDrawingProperties);
                    inline.Append(graphic);

                    drawing.Append(inline);

                    alternateContentChoice.Append(drawing);

                    var alternateContentFallback = new DF.AlternateContentFallback();
                    alternateContentFallback.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

                    alternateContent.Append(alternateContentChoice);
                    alternateContent.Append(alternateContentFallback);

                    run.Append(runProperties);
                    run.Append(alternateContent);

                    paragraph.Append(run);

                    return paragraph;
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Create Shape");
                    throw new Openize.Words.OpenizeException(errorMessage, ex);
                }
            }
        }

        internal static A.GradientFill CreateGradientFill(string color1, string color2)
        {
            A.GradientFill gradientFill = new A.GradientFill();
            A.GradientStopList gradientStopList = new A.GradientStopList();

            // Create start position stop (0%)
            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };
            //A.RgbColorModelHex rgbColor1 = new A.RgbColorModelHex() { Val = "FF0000" };
            A.RgbColorModelHex rgbColor1 = new A.RgbColorModelHex() { Val = color1 };
            gradientStop1.Append(rgbColor1);

            // Create end position stop (100%)
            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 100000 };
            //A.RgbColorModelHex rgbColor2 = new A.RgbColorModelHex() { Val = "0000FF" };
            A.RgbColorModelHex rgbColor2 = new A.RgbColorModelHex() { Val = color2 };
            gradientStop2.Append(rgbColor2);

            gradientStopList.Append(gradientStop1);
            gradientStopList.Append(gradientStop2);

            A.LinearGradientFill linearGradientFill = new A.LinearGradientFill() { Angle = 5400000 }; // 90 degrees

            gradientFill.Append(gradientStopList);
            gradientFill.Append(linearGradientFill);

            return gradientFill;
        }

        internal static A.PatternFill CreatePatternFill(string color1, string color2)
        {
            A.PatternFill patternFill = new A.PatternFill() { Preset = A.PresetPatternValues.SmallGrid };

            A.ForegroundColor fgColor = new A.ForegroundColor();
            //A.RgbColorModelHex fgRgbColor = new A.RgbColorModelHex() { Val = "FF0000" };
            A.RgbColorModelHex fgRgbColor = new A.RgbColorModelHex() { Val = color1 };
            fgColor.Append(fgRgbColor);

            A.BackgroundColor bgColor = new A.BackgroundColor();
            //A.RgbColorModelHex bgRgbColor = new A.RgbColorModelHex() { Val = "FFFFFF" };
            A.RgbColorModelHex bgRgbColor = new A.RgbColorModelHex() { Val = color2 };
            bgColor.Append(bgRgbColor);

            patternFill.Append(fgColor);
            patternFill.Append(bgColor);

            return patternFill;
        }

        internal static A.ShapeTypeValues CreateShapeType(FF.ShapeType shapeType)
        {
            switch (shapeType)
            {
                case FF.ShapeType.Rectangle:
                    return A.ShapeTypeValues.Rectangle;
                case FF.ShapeType.Triangle:
                    return A.ShapeTypeValues.Triangle;
                case FF.ShapeType.Ellipse:
                    return A.ShapeTypeValues.Ellipse;
                case FF.ShapeType.Diamond:
                    return A.ShapeTypeValues.Diamond;
                case FF.ShapeType.Hexagone:
                    return A.ShapeTypeValues.Hexagon;
                default:
                    return A.ShapeTypeValues.Ellipse;
            }
        }

        internal FF.Shape LoadShape(A.GraphicData graphicData, int sequence)
        {
            if (graphicData.Uri.Value == "http://schemas.microsoft.com/office/word/2010/wordprocessingShape")
            {
                var wordprocessingShape = graphicData.GetFirstChild<DWS.WordprocessingShape>();
                if (wordprocessingShape != null)
                {
                    // Extract position and size from shape properties
                    var shapeProperties = wordprocessingShape.GetFirstChild<DWS.ShapeProperties>();
                    var transform2D = shapeProperties.GetFirstChild<A.Transform2D>();

                    var offset = transform2D.Offset;
                    var extents = transform2D.Extents;

                    int x = (int)(offset.X.Value / 9525); // Convert EMU to points
                    int y = (int)(offset.Y.Value / 9525);
                    int width = (int)(extents.Cx.Value / 9525);
                    int height = (int)(extents.Cy.Value / 9525);

                    // Determine the shape type
                    var presetGeometry = shapeProperties.GetFirstChild<A.PresetGeometry>();
                    var shapeType = FF.ShapeType.Ellipse; // Default

                    shapeType = LoadShapeType(presetGeometry.Preset);

                    var shape = new FF.Shape(x, y, width, height, shapeType);
                    shape.ElementId = sequence;

                    // Return the shape object with extracted data
                    return shape;
                }
            }

            return null;
        }

        internal static FF.ShapeType LoadShapeType(A.ShapeTypeValues shapeType)
        {
            if (shapeType == A.ShapeTypeValues.Rectangle)
                return FF.ShapeType.Rectangle;
            else if (shapeType == A.ShapeTypeValues.Triangle)
                return FF.ShapeType.Triangle;
            else if (shapeType == A.ShapeTypeValues.Ellipse)
                return FF.ShapeType.Ellipse;
            else if (shapeType == A.ShapeTypeValues.Diamond)
                return FF.ShapeType.Diamond;
            else if (shapeType == A.ShapeTypeValues.Hexagon)
                return FF.ShapeType.Hexagone;
            else
                return FF.ShapeType.Ellipse;
        }

    }
}
