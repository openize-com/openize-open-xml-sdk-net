using System;
using System.Collections.Generic;
using System.Text;

namespace Openize.Slides.Common.Enumerations
{
    /// <summary>
    /// Specifies the alignment of text elements.
    /// </summary>
    public enum TextAlignment
    {
        Left,
        Right,
        Center,
        None
    }
    /// <summary>
    /// Specifies the type of styled list
    /// </summary>
    public enum ListType
    {
        Bulleted,
        Numbered
    }
    public enum AnimationType
    {
        None,             // No animation
        Zoom,             // Zoom in or out
        FlyIn,            // Fly into the slide
        Spin,             // Spin in place
        FloatIn,         // FloatIn
        Bounce,          // Bounce
    }

}
