
using DocumentFormat.OpenXml.Presentation;


namespace Openize.Slides.Facade.Animations
{
    internal interface IAnimation
    {
        Timing Generate(string shapeId, int duration);
    }
}
