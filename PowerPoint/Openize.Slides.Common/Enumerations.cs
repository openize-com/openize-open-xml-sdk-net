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
        Fade,             // Fade in or out
        Wipe,             // Wipe across the screen
        Zoom,             // Zoom in or out
        FlyIn,            // Fly into the slide
        FlyOut,           // Fly out of the slide
        Bounce,           // Bounce effect
        Spin,             // Spin in place
        GrowShrink,       // Grow or shrink in size
        Flip,             // Flip horizontally or vertically
        Slide,            // Slide in or out
        Morph,            // Morph between shapes or objects
        Appear,           // Appear suddenly
        Dissolve,         // Dissolve into view
        Split,            // Split apart
        Wheel,            // Wheel animation
        Float,            // Float in or out
        Custom            // Custom animation defined by the user
    }

}
