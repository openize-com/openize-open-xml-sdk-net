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
using Openize.Slides.Facade.Animations;


namespace Openize.Slides.Facade
{

    public class AnimateFacade
    {
        // Properties for ShapeId, Delay, and Duration
        public string ShapeId { get; set; }
        public int Duration { get; set; }
        public string Type { get; set; }

        // Constructor to initialize the properties
        public AnimateFacade(string shapeId = "1", int duration = 0, string type = "fade")
        {
            ShapeId = shapeId;
            Type = type;
            Duration = duration;
        }
        public Timing Animate()
        {
            IAnimation animation = Type switch
            {
                "FloatIn" => new FloatIn(),
                "FlyIn" => new FlyIn(),
                "Zoom" => new Zoom(),
                "Spin" => new Spin(),
                "Bounce" => new Bounce(),
                _ => throw new ArgumentException("Unsupported animation type.")
            };

            return animation.Generate(ShapeId, Duration);
        }



    }


}
