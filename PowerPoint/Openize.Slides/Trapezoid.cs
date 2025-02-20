using Openize.Slides.Common;
using Openize.Slides.Common.Enumerations;
using Openize.Slides.Facade;
using System;
using System.Collections.Generic;

namespace Openize.Slides
{
    /// <summary>
    /// Represents a trapezoid shape within a slide.
    /// </summary>
    public class Trapezoid
    {
        private double _x;
        private double _y;
        private double _width;
        private double _height;
        private TrapezoidFacade _facade;
        private int _shapeIndex;
        private string _backgroundColor = null;
        private AnimationType _animation = AnimationType.None;

        /// <summary>
        /// Gets or sets the X coordinate of the trapezoid shape.
        /// </summary>
        public double X { get => _x; set => _x = value; }

        /// <summary>
        /// Gets or sets the Y coordinate of the trapezoid shape.
        /// </summary>
        public double Y { get => _y; set => _y = value; }

        /// <summary>
        /// Gets or sets the width of the trapezoid shape.
        /// </summary>
        public double Width { get => _width; set => _width = value; }

        /// <summary>
        /// Gets or sets the height of the trapezoid shape.
        /// </summary>
        public double Height { get => _height; set => _height = value; }

        /// <summary>
        /// Gets or sets the facade that handles trapezoid shape operations.
        /// </summary>
        public TrapezoidFacade Facade { get => _facade; set => _facade = value; }

        /// <summary>
        /// Gets or sets the shape index within a slide.
        /// </summary>
        public int ShapeIndex { get => _shapeIndex; set => _shapeIndex = value; }

        /// <summary>
        /// Gets or sets the background color of the trapezoid shape.
        /// Default value is "Transparent".
        /// </summary>
        public string BackgroundColor { get => _backgroundColor; set => _backgroundColor = value; }

        /// <summary>
        /// Gets or sets the animation type applied to the trapezoid shape.
        /// Default value is <see cref="AnimationType.None"/>.
        /// </summary>
        public AnimationType Animation { get => _animation; set => _animation = value; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Trapezoid"/> class.
        /// </summary>
        public Trapezoid()
        {
            _facade = new TrapezoidFacade
            {
                ShapeIndex = _shapeIndex
            };

            _backgroundColor = "Transparent";
            _x = Utility.EmuToPixels(1349828);
            _y = Utility.EmuToPixels(1999619);
            _width = Utility.EmuToPixels(6000000);
            _height = Utility.EmuToPixels(2000000);

            PopulateFacade();
        }

        /// <summary>
        /// Updates the trapezoid shape by synchronizing its properties with the facade.
        /// </summary>
        public void Update()
        {
            PopulateFacade();
            _facade.UpdateShape();
        }

        /// <summary>
        /// Populates the facade with the current trapezoid properties.
        /// Converts pixel values to EMUs before setting them in the facade.
        /// </summary>
        private void PopulateFacade()
        {
            _facade.BackgroundColor = _backgroundColor;
            _facade.X = Utility.PixelsToEmu(_x);
            _facade.Y = Utility.PixelsToEmu(_y);
            _facade.Width = Utility.PixelsToEmu(_width);
            _facade.Height = Utility.PixelsToEmu(_height);
        }

        /// <summary>
        /// Retrieves a list of trapezoid objects from their corresponding facades.
        /// </summary>
        /// <param name="trapezoidFacades">A list of <see cref="TrapezoidFacade"/> objects.</param>
        /// <returns>A list of <see cref="Trapezoid"/> objects.</returns>
        public static List<Trapezoid> GetTrapezoids(List<TrapezoidFacade> trapezoidFacades)
        {
            var trapezoids = new List<Trapezoid>();

            try
            {
                foreach (var facade in trapezoidFacades)
                {
                    var trapezoid = new Trapezoid
                    {
                        BackgroundColor = facade.BackgroundColor,
                        X = Utility.EmuToPixels(facade.X),
                        Y = Utility.EmuToPixels(facade.Y),
                        Width = Utility.EmuToPixels(facade.Width),
                        Height = Utility.EmuToPixels(facade.Height),
                        Facade = facade,
                        ShapeIndex = facade.ShapeIndex
                    };

                    trapezoids.Add(trapezoid);
                }
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Getting Trapezoid Shapes");
                throw new Common.OpenizeException(errorMessage, ex);
            }

            return trapezoids;
        }

        /// <summary>
        /// Removes the trapezoid shape from the slide.
        /// </summary>
        public void Remove()
        {
            _facade.RemoveShape(Facade.Trapezoid);
        }
    }
}
