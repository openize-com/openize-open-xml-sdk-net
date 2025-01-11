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
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using Openize.Slides.Common.Enumerations;
using Openize.Slides.Common;
using System.Collections.Generic;
using System;
using DocumentFormat.OpenXml.Bibliography;

namespace Openize.Slides.Facade
{
    public class CommentAuthorFacade
    {
        private String _Name;
        private int _Index;
        private String _InitialLetter;
        private int _ColorIndex;

        public string Name { get => _Name; set => _Name = value; }
        public int Index { get => _Index; set => _Index = value; }
        public string InitialLetter { get => _InitialLetter; set => _InitialLetter = value; }
        public int ColorIndex { get => _ColorIndex; set => _ColorIndex = value; }
        public CommentAuthorFacade() { }

        public void CreateAuthor( ref CommentAuthorList commentAuthorList1) {

            UInt32Value id = new UInt32Value { Value = (uint)_Index };
            UInt32Value color_index = new UInt32Value { Value = (uint)_ColorIndex };
            P.CommentAuthor commentAuthor1 = new P.CommentAuthor() { Id = id, Name = _Name, Initials = _InitialLetter, LastIndex = id, ColorIndex = color_index };

            CommentAuthorExtensionList commentAuthorExtensionList1 = new CommentAuthorExtensionList();

            CommentAuthorExtension commentAuthorExtension1 = new CommentAuthorExtension() { Uri = "{19B8F6BF-5375-455C-9EA6-DF929625EA0E}" };

            P15.PresenceInfo presenceInfo1 = new P15.PresenceInfo() { UserId = _Name, ProviderId = "None" };
            presenceInfo1.AddNamespaceDeclaration("p15", "http://schemas.microsoft.com/office/powerpoint/2012/main");

            commentAuthorExtension1.Append(presenceInfo1);

            commentAuthorExtensionList1.Append(commentAuthorExtension1);

            commentAuthor1.Append(commentAuthorExtensionList1);

            commentAuthorList1.Append(commentAuthor1);
        }
    }
}
