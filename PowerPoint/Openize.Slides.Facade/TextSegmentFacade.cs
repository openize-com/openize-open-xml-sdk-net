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
    public class TextSegmentFacade : Run
    {
        
        private int _FontSize;
        private bool _Bold;
        private bool _Italic;
        private string _FontFamily;
        private string _Color;
        private string _Text;

        public int FontSize
        {
            get
            {
                // Convert fontSizeTwentieths to points
                return (int)_FontSize / 20 / 2;
            }
            set
            {
                // Convert points to twentieths of a point and store in fontSizeTwentieths
                _FontSize = (int)(value * 20 * 2);
            }
        }
        public bool Bold { get => _Bold; set => _Bold = value; }
        public bool Italic { get => _Italic; set => _Italic = value; }
        public string FontFamily { get => _FontFamily; set => _FontFamily = value; }
        public string Color { get => _Color; set => _Color = value; }
        public string Text { get => _Text; set => _Text = value; }

        public TextSegmentFacade ()
        {            
            
        } 
        public void createTextSegment ()
        {
            base.Append(new RunProperties(new SolidFill(new RgbColorModelHex() { Val = _Color }), new LatinFont() { Typeface = _FontFamily })
            { FontSize = _FontSize, Bold = this.Bold, Italic = this.Italic });
            base.Append(new Text() { Text = Text });

        }

        
    }
}
