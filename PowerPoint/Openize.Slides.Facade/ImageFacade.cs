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

namespace Openize.Slides.Facade
{
    public class ImageFacade
    {
        private string _ImagePath;
        private SlidePart _AssociatedSlidePart;
        private ImagePart _PicturePart;
        private string _RelationshipId;
        private int _ImageIndex;
        private P.Picture _Image;
        private Int64Value _x;
        private Int64Value _y;
        private Int64Value _width;
        private Int64Value _height;
        private List<ImageFacade> Images;
        private AnimationType _Animation = AnimationType.None;
        public string ImagePath { get => _ImagePath; set => _ImagePath = value; }
        public SlidePart ImageSlidePart { get => _AssociatedSlidePart; set => _AssociatedSlidePart = value; }
        public ImagePart PicturePart { get => _PicturePart; set => _PicturePart = value; }
        public string RelationshipId { get => _RelationshipId; set => _RelationshipId = value; }
        public int ImageIndex { get => _ImageIndex; set => _ImageIndex = value; }
        public P.Picture Image { get => _Image; set => _Image = value; }
        public Int64Value X { get => _x; set => _x = value; }
        public Int64Value Y { get => _y; set => _y = value; }
        public Int64Value Width { get => _width; set => _width = value; }
        public Int64Value Height { get => _height; set => _height = value; }
        public List<ImageFacade> Images1 { get => Images; set => Images = value; }
        public AnimationType Animation { get => _Animation; set => _Animation = value; }

        public ImageFacade ()
        {
            
        }
        public void createImage (string imagePath, SlidePart slidePart)
        {
            _AssociatedSlidePart = slidePart;
            // Create a unique relationship ID for the image
            _RelationshipId = Utility.GetUniqueRelationshipId();

            // Add the image part to the slide
            _PicturePart = slidePart.AddImagePart(ImagePartType.Png, _RelationshipId);

            using (System.IO.Stream imageStream = new System.IO.FileStream(imagePath, System.IO.FileMode.Open))
            {
                _PicturePart.FeedData(imageStream);
                imageStream.Close();
            }
            _Image = GeneratePicture();

        }
        private  P.Picture GeneratePicture ()
        {
            P.Picture picture1 = new P.Picture();

            P.NonVisualPictureProperties nonVisualPictureProperties1 = new P.NonVisualPictureProperties();

            P.NonVisualDrawingProperties nonVisualDrawingProperties1 = new P.NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Picture 6", Description = "A person in a black and white photo\n\nDescription automatically generated" };

            D.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new D.NonVisualDrawingPropertiesExtensionList();

            D.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new D.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement1 = new OpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{4C9EFE4F-2DDB-7B29-01C3-5A94A8DA92D4}\" />");

            nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

            nonVisualDrawingProperties1.Append(nonVisualDrawingPropertiesExtensionList1);

            P.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new P.NonVisualPictureDrawingProperties();
            D.PictureLocks pictureLocks1 = new D.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
            nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties1);

            P.BlipFill blipFill1 = new P.BlipFill();
            D.Blip blip1 = new D.Blip() { Embed = _RelationshipId };

            D.Stretch stretch1 = new D.Stretch();
            D.FillRectangle fillRectangle1 = new D.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            P.ShapeProperties shapeProperties1 = new P.ShapeProperties();

            D.Transform2D transform2D1 = new D.Transform2D();
            D.Offset offset1 = new D.Offset() { X = _x , Y = _y };
            D.Extents extents1 = new D.Extents() { Cx =_width , Cy =_height  };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            D.PresetGeometry presetGeometry1 = new D.PresetGeometry() { Preset = D.ShapeTypeValues.Rectangle };
            D.AdjustValueList adjustValueList1 = new D.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);
            return picture1;
        }
        public static List<ImageFacade> PopulateImages (SlidePart slidePart)
        {
            var images = new List<ImageFacade>();
            IEnumerable<P.Picture> pictures = slidePart.Slide.CommonSlideData.ShapeTree.Elements<P.Picture>();
            var imageIndex = 0;
            foreach (var pic in pictures)
            {
                
                var imageFacade = new ImageFacade
                {
                    _Image = pic, // Store the P.Shape in the private field
                    
                    X = GetXFromShape(pic),
                    Y = GetYFromShape(pic),
                    Width = GetWidthFromShape(pic),
                    Height = GetHeightFromShape(pic),
                    _AssociatedSlidePart = slidePart,
                    _ImageIndex = imageIndex
                };

                images.Add(imageFacade);
                imageIndex += 1;
            }

            return images;

        }
        private static long GetXFromShape (P.Picture shape)
        {
            return shape.ShapeProperties?.Transform2D?.Offset?.X ?? 0;
        }

        private static long GetYFromShape (P.Picture shape)
        {
            return shape.ShapeProperties?.Transform2D?.Offset?.Y ?? 0;
        }

        private static long GetWidthFromShape (P.Picture shape)
        {
            return shape.ShapeProperties?.Transform2D?.Extents?.Cx ?? 0;
        }

        private static long GetHeightFromShape (P.Picture shape)
        {
            return shape.ShapeProperties?.Transform2D?.Extents?.Cy ?? 0;
        }

        public void RemoveImage (P.Picture image)
        {
            image.Remove();
        }

        public void UpdateImage ()
        {

            if (Image == null)
            {
                throw new InvalidOperationException("Shape has not been created yet. Call CreateShape() first.");
            }
            Image.ShapeProperties.Transform2D = new D.Transform2D()
            {
                Offset = new D.Offset() { X = X, Y = Y },
                Extents = new D.Extents() { Cx = Width, Cy = Height }
            };

           
        }


    }
}
