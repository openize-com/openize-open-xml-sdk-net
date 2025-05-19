using System;
using System.IO;
using DF = DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using FF = Openize.Words.IElements;
using OWD = OpenXML.Words.Data;

namespace OpenXML.Words
{
    internal class OoxmlImage
    {
        private MainDocumentPart _mainPart;
        private readonly object _lockObject = new object();
        private OoxmlImage(MainDocumentPart mainPart)
        {
            _mainPart = mainPart;
        }

        internal static OoxmlImage CreateInstance(MainDocumentPart mainPart)
        {
            return new OoxmlImage(mainPart);
        }

        internal WP.Paragraph CreateImage(FF.Image ffImg)
        {
            lock (_lockObject)
            {
                try
                {
                    var imageBytes = ffImg.ImageData;
                    var imagePart = _mainPart.AddImagePart(ImagePartType.Png);
                    using (var partStream = imagePart.GetStream())
                    {
                        partStream.Write(imageBytes, 0, imageBytes.Length); // Write the image bytes to the partStream
                    }

                    float dpi = 96; // The DPI of the image (you may need to adjust this value)
                    //int widthInPixels;
                    //int heightInPixels;
                    const int maxDimension = 500;

                    var widthInPixels = (ffImg.Width > 0 && ffImg.Width <= maxDimension) ? ffImg.Width : maxDimension;
                    var heightInPixels = (ffImg.Height > 0 && ffImg.Height <= maxDimension) ? ffImg.Height : maxDimension;

                    var widthInInches = widthInPixels / dpi;
                    var heightInInches = heightInPixels / dpi;

                    var widthInEmu = (long)(widthInInches * 914400);
                    var heightInEmu = (long)(heightInInches * 914400);
                    //long widthInEMU = (long)widthInInches;
                    //long heightInEMU = (long)heightInInches;

                    // Define the reference of the image.
                    var element =
                        new WP.Drawing(
                            new DW.Inline(
                                //new DW.Extent() { Cx = ffIMG.Width*9525 , Cy = ffIMG.Height*9525 },
                                new DW.Extent() { Cx = widthInEmu, Cy = heightInEmu },
                                new DW.EffectExtent()
                                {
                                    LeftEdge = 0L,
                                    TopEdge = 0L,
                                    RightEdge = 0L,
                                    BottomEdge = 0L
                                },
                                new DW.DocProperties()
                                {
                                    Id = (DF.UInt32Value)1U,
                                    Name = "Picture 1"
                                },
                                new DW.NonVisualGraphicFrameDrawingProperties(
                                    new A.GraphicFrameLocks() { NoChangeAspect = true }),
                                new A.Graphic(
                                    new A.GraphicData(
                                            new PIC.Picture(
                                                new PIC.NonVisualPictureProperties(
                                                    new PIC.NonVisualDrawingProperties()
                                                    {
                                                        Id = (DF.UInt32Value)0U,
                                                        Name = "New Bitmap Image.jpg"
                                                    },
                                                    new PIC.NonVisualPictureDrawingProperties()),
                                                new PIC.BlipFill(
                                                    new A.Blip(
                                                        new A.BlipExtensionList(
                                                            new A.BlipExtension()
                                                            {
                                                                Uri =
                                                                    "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                            })
                                                    )
                                                    {
                                                        Embed = _mainPart.GetIdOfPart(imagePart),
                                                        CompressionState =
                                                            A.BlipCompressionValues.Print
                                                    },
                                                    new A.Stretch(
                                                        new A.FillRectangle())),
                                                new PIC.ShapeProperties(
                                                    new A.Transform2D(
                                                        new A.Offset() { X = 0L, Y = 0L },
                                                        //new A.Extents() { Cx = ffIMG.Width*9525, Cy = ffIMG.Height*9525 }
                                                        new A.Extents() { Cx = widthInEmu, Cy = heightInEmu }),
                                                    new A.PresetGeometry(
                                                            new A.AdjustValueList()
                                                        )
                                                    { Preset = A.ShapeTypeValues.Rectangle }))
                                        )
                                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                            )
                            {
                                DistanceFromTop = (DF.UInt32Value)0U,
                                DistanceFromBottom = (DF.UInt32Value)0U,
                                DistanceFromLeft = (DF.UInt32Value)0U,
                                DistanceFromRight = (DF.UInt32Value)0U,
                                EditId = "50D07946"
                            });
                    return new WP.Paragraph(new WP.Run(element));
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Create Image");
                    throw new Openize.Words.OpenizeException(errorMessage, ex);
                }
            }
        }

        internal FF.Image LoadImage(WP.Drawing drawing, int sequence)
        {
            lock (_lockObject)
            {
                try
                {
                    foreach (var blip in drawing.Descendants<A.Blip>())
                    {
                        if (blip != null)
                        {
                            var extent = drawing.Inline.Extent;

                            if (extent != null)
                            {
                                var dpi = 96; // Replace with your image's DPI

                                var widthInPixels = (int)(extent.Cx / (914400 / dpi));
                                var heightInPixels = (int)(extent.Cy / (914400 / dpi));

                                var imagePart = _mainPart.GetPartById(blip.Embed) as ImagePart;
                                if (imagePart == null) continue;
                                using var stream = imagePart.GetStream();
                                var image = new FF.Image
                                {
                                    ElementId = sequence
                                };
                                byte[] imageBytes;
                                using (var memoryStream = new MemoryStream())
                                {
                                    stream.CopyTo(memoryStream);
                                    imageBytes = memoryStream.ToArray();
                                }

                                image.ImageData = imageBytes;

                                image.Height = heightInPixels;
                                image.Width = widthInPixels;
                                return image;
                            }
                        }

                    }

                    return null;
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Load Image");
                    throw new Openize.Words.OpenizeException(errorMessage, ex);
                }
            }
        }
    }
}
