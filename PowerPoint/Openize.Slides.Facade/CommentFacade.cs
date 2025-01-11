using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using Openize.Slides.Common;
using System;


namespace Openize.Slides.Facade
{
    public class CommentFacade
    {
        private SlidePart _associatedSlidePart;
        private CommentAuthorsPart _authorsPart;
        private SlideCommentsPart _CommentPart;
        private string _relationshipId;

        

        public SlidePart AssociatedSlidePart { get => _associatedSlidePart; }
        
        public string RelationshipId { get => _relationshipId; set => _relationshipId = value; }
        public CommentAuthorsPart AuthorsPart { get => _authorsPart; set => _authorsPart = value; }
        public SlideCommentsPart CommentPart { get => _CommentPart; set => _CommentPart = value; }

        public CommentFacade()
        {
           
        }
        public void SetAssociatedSlidePart(SlidePart slidePart, CommentAuthorsPart authorPart)
        {
            // Initialize objects
            var commentsPart= slidePart.SlideCommentsPart;
            _associatedSlidePart = slidePart;
            if (slidePart.SlideCommentsPart == null)
            {
                _relationshipId = Utility.GetUniqueRelationshipId();
                commentsPart = _associatedSlidePart.AddNewPart<SlideCommentsPart>(_relationshipId);
            }
            AuthorsPart = authorPart;
            _CommentPart = commentsPart;

            if(_CommentPart.CommentList == null)
            _CommentPart.CommentList=new CommentList();
           
        }

        public void GenerateComment( UInt32Value authorId,  string text, DateTime dateTime, long x, long y)
        {
            var comment = new P.Comment()
            {
                AuthorId = authorId,
                DateTime = System.Xml.XmlConvert.ToDateTime(dateTime.ToString("o"), System.Xml.XmlDateTimeSerializationMode.RoundtripKind),
                Index = new UInt32Value { Value = (uint)_CommentPart.CommentList.Count() + 1 }
            };

            P.Position position = new P.Position() { X = x, Y = y };
            var commentText = new P.Text() { Text = text };

            var commentExtensionList = new CommentExtensionList();
            var commentExtension = new CommentExtension() { Uri = "{C676402C-5697-4E1C-873F-D02D1690AC5C}" };
            var threadingInfo = new P15.ThreadingInfo() { TimeZoneBias = -300 };
            threadingInfo.AddNamespaceDeclaration("p15", "http://schemas.microsoft.com/office/powerpoint/2012/main");
            commentExtension.Append(threadingInfo);
            commentExtensionList.Append(commentExtension);

            comment.Append(position);
            comment.Append(commentText);
            comment.Append(commentExtensionList);

            _CommentPart.CommentList.Append(comment);
        }
        
       public void RemoveComment(int id)
       {
            P.Comment comment = _CommentPart.CommentList.Elements<P.Comment>().FirstOrDefault(c => c.Index == id);
            comment.Remove();
        }


    }
}
