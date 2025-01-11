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
    class ParagraphFacade : Paragraph
    {
        private List<TextSegmentFacade> _TextSegments;
        internal List<TextSegmentFacade> TextSegments { get => _TextSegments; set => _TextSegments = value; }
        public ParagraphFacade ()
        {

        }
        public ParagraphFacade (List<TextSegmentFacade> listTextSegments)
        {
            base.Append(listTextSegments);

        }

        
    }
}
