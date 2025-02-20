using Openize.Slides.Common;
using Openize.Slides.Common.Enumerations;
using Openize.Slides.Facade;
using System;
using System.Collections.Generic;

namespace Openize.Slides
{
    /// <summary>
    /// Represents a rectangle shape within a slide.
    /// </summary>
    public class Rectangle
    {
        /// <summary>
        /// Gets or sets the X coordinate of the rectangle shape.
        /// </summary>
        public double X { get; set; }

        /// <summary>
        /// Gets or sets the Y coordinate of the rectangle shape.
        /// </summary>
        public double Y { get; set; }

        /// <summary>
        /// Gets or sets the width of the rectangle shape.
        /// </summary>
        public double Width { get; set; }

        /// <summary>
        /// Gets or sets the height of the rectangle shape.
        /// </summary>
        public double Height { get; set; }

        /// <summary>
        /// Gets or sets the facade that handles rectangle shape operations.
        /// </summary>
        public RectangleShapeFacade Facade { get; set; }

        /// <summary>
        /// Gets or sets the index of the shape within a slide.
        /// </summary>
        public int ShapeIndex { get; set; }

        /// <summary>
        /// Gets or sets the background color of the rectangle shape.
        /// Default value is "Transparent".
        /// </summary>
        public string BackgroundColor { get; set; } = "Transparent";

        /// <summary>
        /// Gets or sets the animation type applied to the rectangle shape.
        /// Default value is <see cref="AnimationType.None"/>.
        /// </summary>
        public AnimationType Animation { get; set; } = AnimationType.None;

        /// <summary>
        /// Initializes a new instance of the <see cref="Rectangle"/> class.
        /// </summary>
        public Rectangle()
        {
            Facade = new RectangleShapeFacade
            {
                ShapeIndex = ShapeIndex
            };

            X = Utility.EmuToPixels(1349828);
            Y = Utility.EmuToPixels(1999619);
            Width = Utility.EmuToPixels(6000000);
            Height = Utility.EmuToPixels(2000000);

            PopulateFacade();
        }

        /// <summary>
        /// Updates the rectangle shape by synchronizing its properties with the facade.
        /// </summary>
        public void Update()
        {
            PopulateFacade();
            Facade.UpdateShape();
        }

        /// <summary>
        /// Populates the facade with the current rectangle properties.
        /// Converts pixel values to EMUs before setting them in the facade.
        /// </summary>
        private void PopulateFacade()
        {
            Facade.BackgroundColor = BackgroundColor;
            Facade.X = Utility.PixelsToEmu(X);
            Facade.Y = Utility.PixelsToEmu(Y);
            Facade.Width = Utility.PixelsToEmu(Width);
            Facade.Height = Utility.PixelsToEmu(Height);
        }

        /// <summary>
        /// Retrieves a list of rectangle objects from their corresponding facades.
        /// </summary>
        /// <param name="rectangleFacades">A list of <see cref="RectangleShapeFacade"/> objects.</param>
        /// <returns>A list of <see cref="Rectangle"/> objects.</returns>
        public static List<Rectangle> GetRectangles(List<RectangleShapeFacade> rectangleFacades)
        {
            var rectangles = new List<Rectangle>();
            foreach (var facade in rectangleFacades)
            {
                rectangles.Add(new Rectangle
                {
                    BackgroundColor = facade.BackgroundColor,
                    X = Utility.EmuToPixels(facade.X),
                    Y = Utility.EmuToPixels(facade.Y),
                    Width = Utility.EmuToPixels(facade.Width),
                    Height = Utility.EmuToPixels(facade.Height),
                    Facade = facade,
                    ShapeIndex = facade.ShapeIndex
                });
            }
            return rectangles;
        }

        /// <summary>
        /// Removes the rectangle shape from the slide.
        /// </summary>
        public void Remove()
        {
            Facade.RemoveShape(Facade.RectangleShape);
        }
    }
}
