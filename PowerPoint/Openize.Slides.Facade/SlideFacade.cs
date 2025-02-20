using System;
using System.Collections.Generic;
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
namespace Openize.Slides.Facade
{
    public class SlideFacade
    {

        private P.Slide _PresentationSlide = null;
        private SlidePart _SlidePart = null;

        private String _RelationshipId;
        private int _SlideIndex;
        private String _BackgroundColor;

        private List<TextShapeFacade> _TextShapeFacades = null;

        private List<RectangleShapeFacade> _RectangleShapeFacades = null;

        private List<TriangleShapeFacade> _TriangleShapeFacades = null;

        private List<DiamondShapeFacade> _DiamondShapeFacades = null;

        private List<LineFacade> _LineFacades = null;

        private List<CurvedLineFacade> _CurvedLineFacades = null;

        private List<ArrowFacade> _ArrowFacades = null;

        private List<DoubleArrowFacade> _DoubleArrowFacades = null;

        private List<DoubleBraceFacade> _DoubleBraceFacades = null;

        private List<DoubleBracketFacade> _DoubleBracketFacades = null;

        private List<PentagonFacade> _PentagonFacades = null;

        private List<HexagonFacade> _HexagonFacades = null;

        private List<TrapezoidFacade> _TrapezoidFacades = null;

        private List<PieFacade> _PieFacades = null;

        private List<CircleShapeFacade> _CircleShapeFacades = null;

        private List<ImageFacade> _ImagesFacade = null;

        private List<TableFacade> _TableFacades = null;

        private CommentAuthorsPart _CommentAuthorPart;

        private SlideCommentsPart _CommentPart;

        private NotesSlidePart _NotesPart;



        public P.Slide PresentationSlide { get => _PresentationSlide; set => _PresentationSlide = value; }
        public string RelationshipId { get => _RelationshipId; set => _RelationshipId = value; }
        public SlidePart SlidePart { get => _SlidePart; set => _SlidePart = value; }
        public List<TextShapeFacade> TextShapeFacades { get => _TextShapeFacades; set => _TextShapeFacades = value; }
        public int SlideIndex { get => _SlideIndex; set => _SlideIndex = value; }
        public List<ImageFacade> ImagesFacade { get => _ImagesFacade; set => _ImagesFacade = value; }
        public String BackgroundColor { get => _BackgroundColor; set => _BackgroundColor = value; }
        public List<TableFacade> TableFacades { get => _TableFacades; set => _TableFacades = value; }
        public CommentAuthorsPart CommentAuthorPart { get => _CommentAuthorPart; set => _CommentAuthorPart = value; }
        public SlideCommentsPart CommentPart { get => _CommentPart; set => _CommentPart = value; }
        public NotesSlidePart NotesPart { get => _NotesPart; set => _NotesPart = value; }
        public List<RectangleShapeFacade> RectangleShapeFacades { get => _RectangleShapeFacades; set => _RectangleShapeFacades = value; }
        public List<CircleShapeFacade> CircleShapeFacades { get => _CircleShapeFacades; set => _CircleShapeFacades = value; }
        public List<TriangleShapeFacade> TriangleShapeFacades { get => _TriangleShapeFacades; set => _TriangleShapeFacades = value; }
        public List<DiamondShapeFacade> DiamondShapeFacades { get => _DiamondShapeFacades; set => _DiamondShapeFacades = value; }
        public List<LineFacade> LineFacades { get => _LineFacades; set => _LineFacades = value; }
        public List<ArrowFacade> ArrowFacades { get => _ArrowFacades; set => _ArrowFacades = value; }
        public List<DoubleArrowFacade> DoubleArrowFacades { get => _DoubleArrowFacades; set => _DoubleArrowFacades = value; }
        public List<CurvedLineFacade> CurvedLineFacades { get => _CurvedLineFacades; set => _CurvedLineFacades = value; }
        public List<DoubleBraceFacade> DoubleBraceFacades { get => _DoubleBraceFacades; set => _DoubleBraceFacades = value; }
        public List<PentagonFacade> PentagonFacades { get => _PentagonFacades; set => _PentagonFacades = value; }
        public List<DoubleBracketFacade> DoubleBracketFacades { get => _DoubleBracketFacades; set => _DoubleBracketFacades = value; }
        public List<HexagonFacade> HexagonFacades { get => _HexagonFacades; set => _HexagonFacades = value; }
        public List<TrapezoidFacade> TrapezoidFacades { get => _TrapezoidFacades; set => _TrapezoidFacades = value; }
        public List<PieFacade> PieFacades { get => _PieFacades; set => _PieFacades = value; }


        public SlideFacade(bool isNewSlide)
        {
            if (isNewSlide)
            {
                Utility.NextIndex += 1;
                _RelationshipId = Utility.GetUniqueRelationshipId();
                var _PresentationPart = PresentationDocumentFacade.getInstance().GetPresentationPart();
                var _SlideIdList = _PresentationPart.Presentation.SlideIdList;
                _SlideIdList.Append(new SlideId() { Id = (UInt32Value)Utility.GetRandomSlideId(), RelationshipId = _RelationshipId });

                _PresentationSlide = new P.Slide(
                    new CommonSlideData(
                         new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new TransformGroup()))),
                    new ColorMapOverride(new MasterColorMapping()));
                _SlidePart = _PresentationPart.AddNewPart<SlidePart>(_RelationshipId);

                if (PresentationDocumentFacade.getInstance().PresentationSlideLayoutParts.Count != 0)
                    _SlidePart.AddPart(PresentationDocumentFacade.getInstance().PresentationSlideLayoutParts[0]);
            }
        }
        public void SetSlideBackground(string color)
        {
            // Check if color is not null or empty
            if (!string.IsNullOrEmpty(color))
            {
                // Check if there is already a Background element
                Background existingBackground = _PresentationSlide.CommonSlideData.Elements<Background>().FirstOrDefault();

                // If an existing background is found, remove it
                if (existingBackground != null)
                {
                    _PresentationSlide.CommonSlideData.RemoveChild(existingBackground);
                }

                // Create a new Background element with the specified color
                Background newBackground = new Background();
                BackgroundProperties backgroundProperties = new BackgroundProperties();
                SolidFill solidFill = new SolidFill();
                RgbColorModelHex rgbColorModelHex = new RgbColorModelHex() { Val = color };
                solidFill.Append(rgbColorModelHex);
                backgroundProperties.Append(solidFill);
                newBackground.Append(backgroundProperties);

                // Insert the new Background element before the ShapeTree
                _PresentationSlide.CommonSlideData.InsertBefore(newBackground, _PresentationSlide.CommonSlideData.ShapeTree);
            }
        }

        public IEnumerable<Dictionary<string, string>> GetComments()
        {
            List<Dictionary<string, string>> comments = new List<Dictionary<string, string>>();

            if (_CommentPart != null)
            {
                var commentList = _CommentPart.CommentList;
                // Extract comment authors
                foreach (var comment in commentList.Elements<P.Comment>())
                {
                    Dictionary<string, string> CommentProperties = new Dictionary<string, string>
                    {
                        { "Index", comment.Index},
                        { "Text", comment.InnerText },
                        { "AuthorId", comment.AuthorId },
                        { "DateTime", comment.DateTime },
                        { "XPos", comment.Descendants<P.Position>().FirstOrDefault()?.X ?? 0},
                        { "YPos", comment.Descendants<P.Position>().FirstOrDefault()?.X ?? 0 }
                    };

                    comments.Add(CommentProperties);
                }
            }
            return comments;
        }
        public TextShapeFacade AddTextShape(String text, Int32 fontSize, TextAlignment alignment, Int64 _x, Int64 _y,
            Int64 width, Int64 height, String fontFamily, String textColor, String backgroundColor)
        {
            // Create an instance of TextShapeFacade
            TextShapeFacade textShapeFacade = new TextShapeFacade();

            // Set properties using the provided parameters
            textShapeFacade
                .WithText(text)
                .WithFontSize(fontSize)
                .WithFontFamily(fontFamily)
                .WithTextColor(textColor)
                .WithBackgroundColor(backgroundColor)
                .WithAlignment(alignment)
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape textBoxShape = textShapeFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(textBoxShape);
            if (textShapeFacade.Animation != AnimationType.None)
            {
                CallAnimation(textBoxShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id, textShapeFacade.Animation);
            }
            return textShapeFacade;
        }
        public RectangleShapeFacade DrawRectangle(Int64 _x, Int64 _y,
      Int64 width, Int64 height, String backgroundColor, RectangleShapeFacade facade)
        {
            RectangleShapeFacade _RectangleShapeFacade;

            // Use provided facade or create a new one
            if (facade == null)
            {
                _RectangleShapeFacade = new RectangleShapeFacade();
            }
            else
            {
                _RectangleShapeFacade = facade;
            }

            // Set properties
            _RectangleShapeFacade
                .WithBackgroundColor(backgroundColor)
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the shape and append it to the slide
            P.Shape rectangleShape = _RectangleShapeFacade.CreateShape();
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }
            _PresentationSlide.CommonSlideData.ShapeTree.Append(rectangleShape);

            // Handle animation
            if (_RectangleShapeFacade.Animation != AnimationType.None)
            {
                CallAnimation(rectangleShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id, _RectangleShapeFacade.Animation);
            }

            return _RectangleShapeFacade;
        }

        public TriangleShapeFacade DrawTriangle(Int64 _x, Int64 _y,
            Int64 width, Int64 height, String backgroundColor, TriangleShapeFacade facade)
        {
            TriangleShapeFacade _TriangleShapeFacade;

            // Use provided facade or create a new one
            if (facade == null)
            {
                _TriangleShapeFacade = new TriangleShapeFacade();
            }
            else
            {
                _TriangleShapeFacade = facade;
            }

            // Set properties
            _TriangleShapeFacade
                .WithBackgroundColor(backgroundColor)
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the shape and append it to the slide
            P.Shape triangleShape = _TriangleShapeFacade.CreateShape();
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }
            _PresentationSlide.CommonSlideData.ShapeTree.Append(triangleShape);

            // Handle animation
            if (_TriangleShapeFacade.Animation != AnimationType.None)
            {
                CallAnimation(triangleShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id, _TriangleShapeFacade.Animation);
            }

            return _TriangleShapeFacade;
        }

        public DiamondShapeFacade DrawDiamond(Int64 _x, Int64 _y,
            Int64 width, Int64 height, String backgroundColor, DiamondShapeFacade facade)
        {
            DiamondShapeFacade _DiamondShapeFacade;

            // Use provided facade or create a new one
            if (facade == null)
            {
                _DiamondShapeFacade = new DiamondShapeFacade();
            }
            else
            {
                _DiamondShapeFacade = facade;
            }

            // Set properties
            _DiamondShapeFacade
                .WithBackgroundColor(backgroundColor)
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the shape and append it to the slide
            P.Shape diamondShape = _DiamondShapeFacade.CreateShape();
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }
            _PresentationSlide.CommonSlideData.ShapeTree.Append(diamondShape);

            // Handle animation
            if (_DiamondShapeFacade.Animation != AnimationType.None)
            {
                CallAnimation(diamondShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id, _DiamondShapeFacade.Animation);
            }

            return _DiamondShapeFacade;
        }

        public LineFacade DrawLine(Int64 _x, Int64 _y,
            Int64 width, Int64 height, LineFacade facade)
        {
            LineFacade _LineFacade;

            // Use provided facade or create a new one
            if (facade == null)
            {
                _LineFacade = new LineFacade();
            }
            else
            {
                _LineFacade = facade;
            }

            // Set properties
            _LineFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the shape and append it to the slide
            P.Shape line = _LineFacade.CreateShape();
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }
            _PresentationSlide.CommonSlideData.ShapeTree.Append(line);

            // Handle animation
            if (_LineFacade.Animation != AnimationType.None)
            {
                CallAnimation(line.NonVisualShapeProperties.NonVisualDrawingProperties.Id, _LineFacade.Animation);
            }

            return _LineFacade;
        }

        public CurvedLineFacade DrawCurvedLine(Int64 _x, Int64 _y,
            Int64 width, Int64 height, CurvedLineFacade facade)
        {
            CurvedLineFacade _CurvedLineFacade;

            // Use provided facade or create a new one
            if (facade == null)
            {
                _CurvedLineFacade = new CurvedLineFacade();
            }
            else
            {
                _CurvedLineFacade = facade;
            }

            // Set properties
            _CurvedLineFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the shape and append it to the slide
            P.Shape curvedLine = _CurvedLineFacade.CreateShape();
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }
            _PresentationSlide.CommonSlideData.ShapeTree.Append(curvedLine);

            // Handle animation
            if (_CurvedLineFacade.Animation != AnimationType.None)
            {
                CallAnimation(curvedLine.NonVisualShapeProperties.NonVisualDrawingProperties.Id, _CurvedLineFacade.Animation);
            }

            return _CurvedLineFacade;
        }

        public ArrowFacade DrawArrow(Int64 _x, Int64 _y,
            Int64 width, Int64 height, ArrowFacade facade)
        {
            ArrowFacade _ArrowFacade;

            // Use provided facade or create a new one
            if (facade == null)
            {
                _ArrowFacade = new ArrowFacade();
            }
            else
            {
                _ArrowFacade = facade;
            }

            // Set properties
            _ArrowFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the shape and append it to the slide
            P.Shape arrow = _ArrowFacade.CreateShape();
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }
            _PresentationSlide.CommonSlideData.ShapeTree.Append(arrow);

            // Handle animation
            if (_ArrowFacade.Animation != AnimationType.None)
            {
                CallAnimation(arrow.NonVisualShapeProperties.NonVisualDrawingProperties.Id, _ArrowFacade.Animation);
            }

            return _ArrowFacade;
        }

        public DoubleArrowFacade DrawDoubleArrow(Int64 _x, Int64 _y,
            Int64 width, Int64 height, DoubleArrowFacade facade)
        {
            DoubleArrowFacade _DoubleArrowFacade;

            // Use provided facade or create a new one
            if (facade == null)
            {
                _DoubleArrowFacade = new DoubleArrowFacade();
            }
            else
            {
                _DoubleArrowFacade = facade;
            }

            // Set properties
            _DoubleArrowFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the shape and append it to the slide
            P.Shape doubleArrow = _DoubleArrowFacade.CreateShape();
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }
            _PresentationSlide.CommonSlideData.ShapeTree.Append(doubleArrow);

            // Handle animation
            if (_DoubleArrowFacade.Animation != AnimationType.None)
            {
                CallAnimation(doubleArrow.NonVisualShapeProperties.NonVisualDrawingProperties.Id, _DoubleArrowFacade.Animation);
            }

            return _DoubleArrowFacade;
        }

        public DoubleBraceFacade DrawDoubleBrace(Int64 _x, Int64 _y,
            Int64 width, Int64 height, DoubleBraceFacade facade)
        {
            DoubleBraceFacade _DoubleBraceFacade;

            // Use provided facade or create a new one
            if (facade == null)
            {
                _DoubleBraceFacade = new DoubleBraceFacade();
            }
            else
            {
                _DoubleBraceFacade = facade;
            }

            // Set properties
            _DoubleBraceFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the shape and append it to the slide
            P.Shape doubleBrace = _DoubleBraceFacade.CreateShape();
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }
            _PresentationSlide.CommonSlideData.ShapeTree.Append(doubleBrace);

            // Handle animation
            if (_DoubleBraceFacade.Animation != AnimationType.None)
            {
                CallAnimation(doubleBrace.NonVisualShapeProperties.NonVisualDrawingProperties.Id, _DoubleBraceFacade.Animation);
            }

            return _DoubleBraceFacade;
        }

        public DoubleBracketFacade DrawDoubleBracket(Int64 _x, Int64 _y,
            Int64 width, Int64 height, DoubleBracketFacade facade)
        {
            DoubleBracketFacade _DoubleBracketFacade;

            // Use provided facade or create a new one
            if (facade == null)
            {
                _DoubleBracketFacade = new DoubleBracketFacade();
            }
            else
            {
                _DoubleBracketFacade = facade;
            }

            // Set properties
            _DoubleBracketFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the shape and append it to the slide
            P.Shape doubleBracket = _DoubleBracketFacade.CreateShape();
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }
            _PresentationSlide.CommonSlideData.ShapeTree.Append(doubleBracket);

            // Handle animation
            if (_DoubleBracketFacade.Animation != AnimationType.None)
            {
                CallAnimation(doubleBracket.NonVisualShapeProperties.NonVisualDrawingProperties.Id, _DoubleBracketFacade.Animation);
            }

            return _DoubleBracketFacade;
        }

        public PentagonFacade DrawPentagon(Int64 _x, Int64 _y,
     Int64 width, Int64 height, PentagonFacade facade)
        {
            // Use the provided PentagonFacade or create a new instance
            PentagonFacade _PentagonFacade = facade ?? new PentagonFacade();

            // Set properties using the provided parameters
            _PentagonFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape _Pentagon = _PentagonFacade.CreateShape();

            // Append the shape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(_Pentagon);

            // Handle animation if necessary
            if (_PentagonFacade.Animation != AnimationType.None)
            {
                CallAnimation(_Pentagon.NonVisualShapeProperties.NonVisualDrawingProperties.Id, _PentagonFacade.Animation);
            }

            return _PentagonFacade;
        }

        public HexagonFacade DrawHexagon(Int64 _x, Int64 _y,
            Int64 width, Int64 height, HexagonFacade facade)
        {
            // Use the provided HexagonFacade or create a new instance
            HexagonFacade _HexagonFacade = facade ?? new HexagonFacade();

            // Set properties using the provided parameters
            _HexagonFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape _Hexagon = _HexagonFacade.CreateShape();

            // Append the shape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(_Hexagon);

            // Handle animation if necessary
            if (_HexagonFacade.Animation != AnimationType.None)
            {
                CallAnimation(_Hexagon.NonVisualShapeProperties.NonVisualDrawingProperties.Id, _HexagonFacade.Animation);
            }

            return _HexagonFacade;
        }

        public TrapezoidFacade DrawTrapezoid(Int64 _x, Int64 _y,
            Int64 width, Int64 height, TrapezoidFacade facade)
        {
            // Use the provided TrapezoidFacade or create a new instance
            TrapezoidFacade _TrapezoidFacade = facade ?? new TrapezoidFacade();

            // Set properties using the provided parameters
            _TrapezoidFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape _Trapezoid = _TrapezoidFacade.CreateShape();

            // Append the shape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(_Trapezoid);

            // Handle animation if necessary
            if (_TrapezoidFacade.Animation != AnimationType.None)
            {
                CallAnimation(_Trapezoid.NonVisualShapeProperties.NonVisualDrawingProperties.Id, _TrapezoidFacade.Animation);
            }

            return _TrapezoidFacade;
        }

        public CircleShapeFacade DrawCircle(Int64 _x, Int64 _y,
    Int64 width, Int64 height, String backgroundColor, CircleShapeFacade facade)
        {
            CircleShapeFacade _CircleShapeFacade;

            // Create an instance of CircleShapeFacade if facade is null, otherwise use the provided facade
            if (facade == null)
            {
                _CircleShapeFacade = new CircleShapeFacade();
            }
            else
            {
                _CircleShapeFacade = facade;
            }

            // Set properties using the provided parameters
            _CircleShapeFacade
                .WithBackgroundColor(backgroundColor)
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape CircleShape = _CircleShapeFacade.CreateShape();

            // Append the CircleShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(CircleShape);

            // Handle animation if present
            if (_CircleShapeFacade.Animation != AnimationType.None)
            {
                CallAnimation(CircleShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id, _CircleShapeFacade.Animation);
            }

            return _CircleShapeFacade;
        }

        public PieFacade DrawPie(Int64 _x, Int64 _y,
         Int64 width, Int64 height, PieFacade facade)
        {
            PieFacade _PieFacade;
            // Create an instance of TextShapeFacade
            if (facade == null)
            {
                _PieFacade = new PieFacade();
            }
            else
            {
                _PieFacade = facade;
            }
            _PieFacade.Animation = facade.Animation;
            // Set properties using the provided parameters
            _PieFacade
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape _Pie = _PieFacade.CreateShape();

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }
            _PresentationSlide.CommonSlideData.ShapeTree.Append(_Pie);
            if (_PieFacade.Animation != AnimationType.None)
            {
                CallAnimation(_Pie.NonVisualShapeProperties.NonVisualDrawingProperties.Id, _PieFacade.Animation);
            }
            return _PieFacade;
        }
        public TextShapeFacade AddTextListShape(List<String> textList, ListFacade facade, Int32 fontSize, TextAlignment alignment, Int64 _x, Int64 _y,
            Int64 width, Int64 height, String fontFamily, String textColor, String backgroundColor)
        {
            // Create an instance of TextShapeFacade
            TextShapeFacade textShapeFacade = new TextShapeFacade();

            // Set properties using the provided parameters
            textShapeFacade
                .WithFontSize(fontSize)
                .WithFontFamily(fontFamily)
                .WithTextColor(textColor)
                .WithBackgroundColor(backgroundColor)
                .WithAlignment(alignment)
                .WithPosition(_x, _y)
                .WithSize(width, height);

            // Create the P.Shape using the CreateShape method
            P.Shape textBoxShape = textShapeFacade.CreateListShape(textList, facade);

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(textBoxShape);
            //_TextShapeFacades.Add(textShapeFacade);
            if (textShapeFacade.Animation != AnimationType.None)
            {
                CallAnimation(textBoxShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id, textShapeFacade.Animation);
            }

            return textShapeFacade;
        }
        public TextShapeFacade AddTextShape(List<TextSegmentFacade> textSegmentFacades, TextAlignment alignment, Int64 _x, Int64 _y,
           Int64 width, Int64 height, String backgroundColor)
        {
            // Create an instance of TextShapeFacade
            TextShapeFacade textShapeFacade = new TextShapeFacade();

            // Set properties using the provided parameters
            textShapeFacade
                .WithAlignment(alignment)
                .WithPosition(_x, _y)
                .WithSize(width, height)
                .WithBackgroundColor(backgroundColor);

            // Create the P.Shape using the CreateShape method
            P.Shape textBoxShape = textShapeFacade.CreateShape(textSegmentFacades);

            // Append the textBoxShape to the ShapeTree of the presentation slide
            if (_PresentationSlide.CommonSlideData.ShapeTree == null)
            {
                _PresentationSlide.CommonSlideData.ShapeTree = new P.ShapeTree();
            }

            _PresentationSlide.CommonSlideData.ShapeTree.Append(textBoxShape);
            //_TextShapeFacades.Add(textShapeFacade);
            if (textShapeFacade.Animation != AnimationType.None)
            {
                CallAnimation(textBoxShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id, textShapeFacade.Animation);
            }
            return textShapeFacade;
        }

        public void AddNote(String noteText)
        {
            var relId = _RelationshipId;

            NotesSlidePart notesSlidePart1;
            string existingSlideNote = noteText;

            if (_SlidePart.NotesSlidePart != null)
            {
                //Appened new note to existing note.
                existingSlideNote = _SlidePart.NotesSlidePart.NotesSlide.InnerText + "\n" + noteText;
                //var val = (NotesSlidePart)_SlidePart.GetPartById(relId);
                //var val = _SlidePart.NotesSlidePart;
                notesSlidePart1 = _NotesPart;
            }
            else
            {
                //Add a new noteto a slide.                      
                notesSlidePart1 = _SlidePart.AddNewPart<NotesSlidePart>(relId);
            }

            NotesSlide notesSlide = new NotesSlide(
                new CommonSlideData(new ShapeTree(
                  new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new D.TransformGroup()),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Slide Image Placeholder 1" },
                            new P.NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true, NoRotation = true, NoChangeAspect = true }),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.SlideImage })),
                        new P.ShapeProperties()),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Notes Placeholder 2" },
                            new P.NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Body, Index = (UInt32Value)1U })),
                        new P.ShapeProperties(),
                        new P.TextBody(
                            new D.BodyProperties(),
                            new D.ListStyle(),
                            new D.Paragraph(
                                new D.Run(
                                    new D.RunProperties() { Language = "en-US", Dirty = false },
                                    new D.Text() { Text = existingSlideNote }),
                                new D.EndParagraphRunProperties() { Language = "en-US", Dirty = false }))
                            ))),
                new ColorMapOverride(new D.MasterColorMapping()));

            notesSlidePart1.NotesSlide = notesSlide;
        }
        public void RemoveNote()
        {
            if (_SlidePart.NotesSlidePart != null)
            {
                // Clear the existing notes.
                _SlidePart.DeletePart(_SlidePart.NotesSlidePart);
            }
        }
        public void AddImage(ImageFacade picture)
        {
            _PresentationSlide.CommonSlideData.ShapeTree.Append(picture.Image);
        }

        public void Update()
        {
            this.SetSlideBackground(_BackgroundColor);
        }
        private void CallAnimation(string shapeId, AnimationType animation)
        {

            AnimateFacade animateFacade = new AnimateFacade();
            animateFacade.Type = animation.ToString();

            // Optionally, override default properties
            animateFacade.ShapeId = shapeId; // You can change the ShapeId if needed

            _PresentationSlide.Append(animateFacade.animate());
        }
        public void close()
        {
            _PresentationSlide.RemoveAllChildren();
            _PresentationSlide.Remove();
            _PresentationSlide = null;
        }

    }
}
