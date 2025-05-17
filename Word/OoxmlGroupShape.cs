using System;
using System.Collections.Generic;
using DF = DocumentFormat.OpenXml;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DWG = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using DWS = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using FF = Openize.Words.IElements;
using OWD = OpenXML.Words.Data;

namespace OpenXML.Words
{
    internal class OoxmlGroupShape
    {
        private readonly object _lockObject = new object();

        private OoxmlGroupShape()
        {
        }

        internal static OoxmlGroupShape CreateInstance()
        {
            return new OoxmlGroupShape();
        }

        internal WP.Paragraph CreateGroupShape(FF.GroupShape groupShape)
        {
            lock (_lockObject)
            {
                try
                {
                    if (groupShape.Shape2.X < (groupShape.Shape1.X + groupShape.Shape1.Width))
                        throw new Openize.Words.OpenizeException("Invalid shape dimensions",
                            new ArgumentException());

                    var paragraph = new WP.Paragraph();
                    paragraph.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordml");

                    var run = new WP.Run();

                    var runProperties = new WP.RunProperties();
                    var noProof = new WP.NoProof();

                    runProperties.Append(noProof);

                    var alternateContent = new DF.AlternateContent();
                    alternateContent.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

                    var alternateContentChoice = new DF.AlternateContentChoice() { Requires = "wpg" };
                    alternateContentChoice.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

                    var drawing = new WP.Drawing();
                    drawing.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                    var inline = new DW.Inline() { DistanceFromTop = (DF.UInt32Value)0U, DistanceFromBottom = (DF.UInt32Value)0U, DistanceFromLeft = (DF.UInt32Value)0U, DistanceFromRight = (DF.UInt32Value)0U, AnchorId = "24C249F3", EditId = "163BC827" };
                    inline.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
                    inline.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");

                    var extent = new DW.Extent() { Cx = 3778250L, Cy = 622300L };
                    extent.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

                    var effectExtent = new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 12700L, BottomEdge = 25400L };
                    effectExtent.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

                    var docProperties = new DW.DocProperties() { Id = (DF.UInt32Value)122768519U, Name = "Group-" + groupShape.ElementId.ToString() };
                    docProperties.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

                    var nonVisualGraphicFrameDrawingProperties = new DW.NonVisualGraphicFrameDrawingProperties();
                    nonVisualGraphicFrameDrawingProperties.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

                    var graphic = new A.Graphic();
                    graphic.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                    A.GraphicData graphicData = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" };

                    var wordprocessingGroup = new DWG.WordprocessingGroup();
                    wordprocessingGroup.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
                    var nonVisualGroupDrawingShapeProperties = new DWG.NonVisualGroupDrawingShapeProperties();

                    var shape1 = groupShape.Shape1;
                    var shape2 = groupShape.Shape2;

                    var groupShapeProperties = new DWG.GroupShapeProperties();

                    var transformGroup = new A.TransformGroup();

                    /****** Group Offsets ******/
                    // var offset = new A.Offset() { X = 0L, Y = 0L };
                    // Group.X=shape1.x, Group.Y=shape1.y
                    var groupX = shape1.X * 9525;
                    var groupY = shape1.Y * 9525;
                    var groupOffset = new A.Offset()
                    {
                        X = groupX,
                        Y = groupY
                    };

                    /****** Group Extents ******/
                    // var extents = new A.Extents() { Cx = 3778250L, Cy = 622300L };
                    // Group.Width=278(shape2.X)-0(shape1.X)+118(shape2.Width)
                    var groupCx = (shape2.X - shape1.X + shape2.Width) * 9525;
                    // Group.Height=Shape1.Height
                    var groupCy = shape1.Height * 9525;
                    var groupExtents = new A.Extents() { Cx = groupCx, Cy = groupCy };

                    /****** Child Offset & Extents ******/
                    //var childOffset = new A.ChildOffset() { X=0L, Y=0L};
                    // Same as group.X and group.Y
                    var childOffset = new A.ChildOffset()
                    {
                        X = groupX,
                        Y = groupY
                    };
                    //var childExtents = new A.ChildExtents() { Cx = 3778250L, Cy = 622300L };
                    var childExtents = new A.ChildExtents()
                    {
                        Cx = groupCx,
                        Cy = groupCy
                    };

                    transformGroup.Append(groupOffset);
                    transformGroup.Append(groupExtents);
                    transformGroup.Append(childOffset);
                    transformGroup.Append(childExtents);

                    groupShapeProperties.Append(transformGroup);

                    /******************* shapes ****************/
                    var wordprocessingShape01 = CreatePartialShape(
                        shape1.ElementId, shape1.X, shape1.Y,
                        shape1.Width, shape1.Height, OoxmlShape.CreateShapeType(
                            shape1.Type), shape1);
                    var wordprocessingShape02 = CreatePartialShape(
                        shape2.ElementId, shape2.X, shape2.Y,
                        shape2.Width, shape2.Height, OoxmlShape.CreateShapeType(
                            shape2.Type), shape2
                        );

                    /**************** connector *****************/
                    // Connector
                    var wordprocessingShape = new DWS.WordprocessingShape();
                    wordprocessingShape.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
                    var nonVisualDrawingProperties = new DWS.NonVisualDrawingProperties()
                    {
                        Id = (DF.UInt32Value)161453463U,
                        Name = "Connector: Elbow 161453463"
                    };

                    var nonVisualConnectorProperties = new DWS.NonVisualConnectorProperties();
                    //A.StartConnection startConnection1 = new A.StartConnection() { Id = (DF.UInt32Value)448142074U, Index = (DF.UInt32Value)3U };
                    A.StartConnection startConnection = new A.StartConnection()
                    {
                        Id = (DF.UInt32Value)(uint)shape1.ElementId,
                        Index = (DF.UInt32Value)3U
                    };
                    //A.EndConnection endConnection1 = new A.EndConnection() { Id = (DF.UInt32Value)1011268246U, Index = (DF.UInt32Value)2U };
                    A.EndConnection endConnection = new A.EndConnection()
                    {
                        Id = (DF.UInt32Value)(uint)shape2.ElementId,
                        Index = (DF.UInt32Value)2U
                    };

                    nonVisualConnectorProperties.Append(startConnection);
                    nonVisualConnectorProperties.Append(endConnection);

                    var shapeProperties = new DWS.ShapeProperties();

                    var transform2D = new A.Transform2D();
                    //var offset4 = new A.Offset() { X = 914400L, Y = 311150L };
                    // 96 (same as shape1.width),33 (half of shape1.height)
                    var connectorX = shape1.Width * 9525;
                    var connectorY = shape1.Height / 2 * 9525;
                    var offset4 = new A.Offset()
                    {
                        X = connectorX,
                        Y = connectorY
                    };
                    //var extents4 = new A.Extents() { Cx = 1733550L, Cy = 6350L };
                    // 182, 1
                    // connector.Cx = Group.Width - (shape1.Width+shape2.Width)
                    var connectorCx = groupCx - (connectorX + (shape2.Width * 9525));
                    var extents4 = new A.Extents() { Cx = connectorCx, Cy = 6350L };

                    transform2D.Append(offset4);
                    transform2D.Append(extents4);

                    var presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.BentConnector3 };
                    var adjustValueList = new A.AdjustValueList();

                    presetGeometry.Append(adjustValueList);
                    var outline = new A.Outline();

                    shapeProperties.Append(transform2D);
                    shapeProperties.Append(presetGeometry);
                    shapeProperties.Append(outline);

                    var shapeStyle = new DWS.ShapeStyle();

                    var lineReference = new A.LineReference() { Index = (DF.UInt32Value)1U };
                    var schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

                    lineReference.Append(schemeColor);

                    var fillReference = new A.FillReference() { Index = (DF.UInt32Value)0U };
                    schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

                    fillReference.Append(schemeColor);

                    var effectReference = new A.EffectReference() { Index = (DF.UInt32Value)0U };
                    A.RgbColorModelPercentage rgbColorModelPercentage = new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

                    effectReference.Append(rgbColorModelPercentage);

                    var fontReference = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
                    schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

                    fontReference.Append(schemeColor);

                    shapeStyle.Append(lineReference);
                    shapeStyle.Append(fillReference);
                    shapeStyle.Append(effectReference);
                    shapeStyle.Append(fontReference);
                    var textBodyProperties = new DWS.TextBodyProperties();

                    wordprocessingShape.Append(nonVisualDrawingProperties);
                    wordprocessingShape.Append(nonVisualConnectorProperties);
                    wordprocessingShape.Append(shapeProperties);
                    wordprocessingShape.Append(shapeStyle);
                    wordprocessingShape.Append(textBodyProperties);

                    wordprocessingGroup.Append(nonVisualGroupDrawingShapeProperties);
                    wordprocessingGroup.Append(groupShapeProperties);
                    wordprocessingGroup.Append(wordprocessingShape01);
                    wordprocessingGroup.Append(wordprocessingShape02);
                    wordprocessingGroup.Append(wordprocessingShape);

                    graphicData.Append(wordprocessingGroup);

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
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Create Group Shape");
                    throw new Openize.Words.OpenizeException(errorMessage, ex);
                }
            }
        }

        private DWS.WordprocessingShape CreatePartialShape(int Id, int X, int Y, int Width, int Height,
            A.ShapeTypeValues shapeTypeValues,
            FF.Shape shape)
        {
            lock (_lockObject)
            {
                try
                {
                    var wordprocessingShape = new DWS.WordprocessingShape();
                    wordprocessingShape.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

                    //var nonVisualDrawingProperties1 = new DWS.NonVisualDrawingProperties() { Id = (DF.UInt32Value)448142074U, Name = "Rectangle 448142074" };
                    var nonVisualDrawingProperties = new DWS.NonVisualDrawingProperties()
                    {
                        Id = (DF.UInt32Value)(uint)Id,
                        Name = "shape-" + Id.ToString()
                    };
                    var nonVisualDrawingShapeProperties = new DWS.NonVisualDrawingShapeProperties();

                    var shapeProperties = new DWS.ShapeProperties();
                    var transform2D = new A.Transform2D();
                    //var offset = new A.Offset() { X = 0L, Y = 0L };
                    var offset = new A.Offset() { X = X * 9525, Y = Y * 9525 };
                    //var extents = new A.Extents() { Cx = 914400L, Cy = 622300L };
                    var extents = new A.Extents() { Cx = Width * 9525, Cy = Height * 9525 };

                    transform2D.Append(offset);
                    transform2D.Append(extents);

                    //var presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                    var presetGeometry = new A.PresetGeometry()
                    {
                        Preset = shapeTypeValues
                    };
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
                            shapeProperties.Append(OoxmlShape.CreateGradientFill(
                                shape.FillColors.Color1, shape.FillColors.Color2));
                            break;
                        case FF.ShapeFillType.Pattern:
                            shapeProperties.Append(OoxmlShape.CreatePatternFill(
                                shape.FillColors.Color1, shape.FillColors.Color2));
                            break;
                    }
                    shapeProperties.Append(outline);

                    var shapeStyle = new DWS.ShapeStyle();

                    var lineReference = new A.LineReference()
                    {
                        Index = (DF.UInt32Value)2U
                    };

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

                    wordprocessingShape.Append(nonVisualDrawingProperties);
                    wordprocessingShape.Append(nonVisualDrawingShapeProperties);
                    wordprocessingShape.Append(shapeProperties);
                    wordprocessingShape.Append(shapeStyle);
                    wordprocessingShape.Append(textBodyProperties);

                    return wordprocessingShape;
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Create Partial Shape");
                    throw new Openize.Words.OpenizeException(errorMessage, ex);
                }
            }
        }

        internal FF.GroupShape LoadGroupShape(A.GraphicData graphicData, int sequence)
        {
            lock (_lockObject)
            {
                try
                {

                    if (graphicData.Uri.Value == "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup")
                    {
                        var wordprocessingShapes = graphicData.Descendants<DWS.WordprocessingShape>();//.GetFirstChild<DWS.WordprocessingShape>();
                        if (wordprocessingShapes != null)
                        {
                            var count = 0;
                            var listShapes = new List<FF.Shape>();
                            foreach (var wpShape in wordprocessingShapes)
                            {
                                var shapeProperties = wpShape.GetFirstChild<DWS.ShapeProperties>();
                                var presetGeometry = shapeProperties.GetFirstChild<A.PresetGeometry>();
                                var shapeType = FF.ShapeType.Ellipse;

                                if (presetGeometry.Preset != A.ShapeTypeValues.BentConnector3)
                                {
                                    count = count + 1;
                                    var transform2D = shapeProperties.GetFirstChild<A.Transform2D>();

                                    var offset = transform2D.Offset;
                                    var extents = transform2D.Extents;

                                    int x = (int)(offset.X.Value / 9525); // Convert EMU to points
                                    int y = (int)(offset.Y.Value / 9525);
                                    int width = (int)(extents.Cx.Value / 9525);
                                    int height = (int)(extents.Cy.Value / 9525);
                                    shapeType = OoxmlShape.LoadShapeType(presetGeometry.Preset);
                                    var shape = new FF.Shape(x, y, width, height, shapeType);
                                    shape.ElementId = sequence * 50 + count;
                                    listShapes.Add(shape);
                                }
                            }
                            var groupShape = new FF.GroupShape(listShapes[0], listShapes[1]);
                            groupShape.ElementId = sequence;
                            return groupShape;
                        }
                    }
                    return null;
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Load Group Shape");
                    throw new Openize.Words.OpenizeException(errorMessage, ex);
                }
            }
        }
    }
}
