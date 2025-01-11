using Openize.Slides.Common;
using Openize.Slides.Common.Enumerations;
using Openize.Slides.Facade;
using System;
using System.Collections.Generic;

namespace Openize.Slides
{
    /// <summary>
    /// This class represents the image within a slide.
    /// </summary>
    public class Image
    {
        private string _ImagePath;
        private ImageFacade _Facade;
        private int _ImageIndex = 0;
        private String _Name;
        private double _x;
        private double _y;
        private double _Width;
        private double _Height;
        private AnimationType _Animation = AnimationType.None;
        /// <summary>
        /// Property to get or set the image path.
        /// </summary>
        public string ImagePath { get => _ImagePath; set => _ImagePath = value; }
        /// <summary>
        /// Property to get or set the ImageFacade instance.
        /// </summary>
        public ImageFacade Facade { get => _Facade; set => _Facade = value; }
        /// <summary>
        /// Property to get or set the image index within the slide.
        /// </summary>
        public int ImageIndex { get => _ImageIndex; set => _Facade.ImageIndex = _ImageIndex = value; }
        /// <summary>
        /// Property to get or set the image index within the slide.
        /// </summary>
        public string Name { get => _Name; set => _Name = value; }
        /// <summary>
        /// Property to get or set the X coordinate of an image.
        /// </summary>
        public double X { get => _x; set => _x = value; }
        /// <summary>
        /// Property to get or set the Y coordinate of an image.
        /// </summary>
        public double Y { get => _y; set => _y = value; }
        /// <summary>
        /// Property to get or set the width of an image.
        /// </summary>
        public double Width { get => _Width; set => _Width = value; }
        /// <summary>
        /// Property to get or set the height of an image.
        /// </summary>
        public double Height { get => _Height; set => _Height = value; }
        /// <summary>
        /// Property to set animation
        /// </summary>
        public AnimationType Animation { get => _Animation; set => _Animation = value; }
        /// <summary>
        /// Initialize the image object 
        /// </summary>
        /// <param name="imagePath">Image path as string</param>
        public Image (String imagePath)
        {
            _ImagePath = imagePath;


        }
        /// <summary>
        /// Blank constructor to initialize the image object
        /// </summary>
        public Image ()
        {

        }
        /// <summary>
        /// Method to get the list of the images within a slide
        /// </summary>
        /// <param name="imageFacades">An object of ImageFacade.</param>
        /// <returns></returns>
        public static List<Image> GetImages (List<ImageFacade> imageFacades)
        {
            var pictures = new List<Image>();

            foreach (var facade in imageFacades)
            {
                Image pic = new Image
                {
                    X = Utility.EmuToPixels(facade.X),
                    Y = Utility.EmuToPixels(facade.Y),
                    Width = Utility.EmuToPixels(facade.Width),
                    Height = Utility.EmuToPixels(facade.Height),
                    Facade = facade,
                    ImageIndex = facade.ImageIndex
                };

                pictures.Add(pic);

            }

            return pictures;
        }
        public void Update ()
        {
            Populate_Facade();
            _Facade.UpdateImage();

        }
        /// <summary>
        /// Method to remove the image.
        /// </summary>
        public void Remove ()
        {
            _Facade.RemoveImage(this._Facade.Image);
        }
        /// <summary>
        /// Method to populate Facade respective to image.
        /// </summary>
        private void Populate_Facade ()
        {
            _Facade.X = Utility.PixelsToEmu(_x);
            _Facade.Y = Utility.PixelsToEmu(_y);
            _Facade.Width = Utility.PixelsToEmu(_Width);
            _Facade.Height = Utility.PixelsToEmu(_Height);

        }
    }
}
