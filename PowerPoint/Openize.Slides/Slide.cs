using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using Openize.Slides.Common;
using Openize.Slides.Common.Enumerations;
using Openize.Slides.Facade;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Openize.Slides
{
    /// <summary>
    /// Represents the slide object within a presentatction
    /// </summary>
    public class Slide
    {
        private SlideFacade _SlideFacade;
        private String _RelationshipId;
        private int _SlideIndex;
        private List<TextShape> _TextShapes;
        private List<Rectangle> _Rectangles;
        private List<Triangle> _Triangles;
        private List<Diamond> _Diamonds;
        private List<Line> _Lines;
        private List<CurvedLine> _CurvedLines;
        private List<Arrow> _Arrows;
        private List<DoubleArrow> _DoubleArrows;
        private List<DoubleBrace> _DoubleBraces;
        private List<DoubleBracket> _DoubleBrackets;
        private List<Pentagon> _Pentagons;
        private List<Hexagon> _Hexagons;
        private List<Pie> _Pies;
        private List<Trapezoid> _Trapezoids;
        private List<Circle> _Circles;
        private List<Image> _Images;
        private List<Table> _Tables;
        private static String _BackgroundColor = null;
        private CommentFacade _CommentFacade=null;
        private Presentation _SlidePresentation;
        private int _CommentIndex = 0;
        private AnimationType _Animation = AnimationType.None;

        /// <summary>
        /// Property for respective Slide Facade.
        /// </summary>
        public SlideFacade SlideFacade { get => _SlideFacade; set => _SlideFacade = value; }
        /// <summary>
        /// Property contains the list of all text shapes.
        /// </summary>
        public List<TextShape> TextShapes { get => _TextShapes; set => _TextShapes = value; }
        /// <summary>
        /// Property for the relationship Id.
        /// </summary>
        public string RelationshipId { get => _RelationshipId; set => _RelationshipId = value; }

        /// <summary>
        /// Property to hold the index of the slide.
        /// </summary>
        public int SlideIndex { get => _SlideIndex; set => _SlideIndex = value; }
        /// <summary>
        /// Property contains the list of all images within a slide.
        /// </summary>
        public List<Image> Images { get => _Images; set => _Images = value; }
        /// <summary>
        /// Property to set background color of a slide.
        /// </summary>
        public string BackgroundColor { get => _BackgroundColor; set => _BackgroundColor = value; }
        /// <summary>
        /// Property to get or set the list of tables
        /// </summary>
        public List<Table> Tables { get => _Tables; set => _Tables = value; }
        /// <summary>
        /// Property to get or set the relative presentation instance
        /// </summary>
        public Presentation SlidePresentation { get => _SlidePresentation; set => _SlidePresentation = value; }
        /// <summary>
        /// Property to get or set the list of Rectangles.
        /// </summary>
        public List<Rectangle> Rectangles { get => _Rectangles; set => _Rectangles = value; }
        /// <summary>
        /// Property to get or set list of circles.
        /// </summary>
        public List<Circle> Circles { get => _Circles; set => _Circles = value; }
        /// <summary>
        /// Property to get or set list of diamonds.
        /// </summary>
        public List<Diamond> Diamonds { get => _Diamonds; set => _Diamonds = value; }
        /// <summary>
        /// Property to get or set list of triangles.
        /// </summary>
        public List<Triangle> Triangles { get => _Triangles; set => _Triangles = value; }
        /// <summary>
        /// Property to get or set list of lines.
        /// </summary>
        public List<Line> Lines { get => _Lines; set => _Lines = value; }
        /// <summary>
        /// Property to get or set list of arrows.
        /// </summary>
        public List<Arrow> Arrows { get => _Arrows; set => _Arrows = value; }
        /// <summary>
        /// Property to get or set list of double arrows.
        /// </summary>
        public List<DoubleArrow> DoubleArrows { get => _DoubleArrows; set => _DoubleArrows = value; }
        /// <summary>
        /// Property to get or set list of curved lines.
        /// </summary>
        public List<CurvedLine> CurvedLines { get => _CurvedLines; set => _CurvedLines = value; }

        /// <summary>
        /// Property to get or set list of double braces.
        /// </summary>
        public List<DoubleBrace> DoubleBraces { get => _DoubleBraces; set => _DoubleBraces = value; }

        /// <summary>
        /// Property to get or set list of Pentagons.
        /// </summary>
        public List<Pentagon> Pentagons { get => _Pentagons; set => _Pentagons = value; }

        /// <summary>
        /// Property to get or set list of double bracket.
        /// </summary>
        public List<DoubleBracket> DoubleBrackets { get => _DoubleBrackets; set => _DoubleBrackets = value; }

        /// <summary>
        /// Property to get or set list of Hexagon.
        /// </summary>
        public List<Hexagon> Hexagons { get => _Hexagons; set => _Hexagons = value; }

        /// <summary>
        /// Property to get or set list of Trapezoid.
        /// </summary>
        public List<Trapezoid> Trapezoids { get => _Trapezoids; set => _Trapezoids = value; }
        /// <summary>
        /// Property to get or set list of Pie.
        /// </summary>
        public List<Pie> Pies { get => _Pies; set => _Pies = value; }

       

        /// <summary>
        /// Constructor for the Slide class.
        /// </summary>
        /// <remarks>
        ///  it intializes the Slide Facade set the slide index and intializes the lists of text shapes and images.
        /// </remarks>
        public Slide ()
        {
            try
            {
                _SlideIndex = Utility.SlideNextIndex;
                Utility.SlideNextIndex += 1;
                _SlideFacade = new SlideFacade(true);
                _SlideFacade.SlideIndex = _SlideIndex;
                _RelationshipId = _SlideFacade.RelationshipId;
                _TextShapes = new List<TextShape>();
                _Rectangles = new List<Rectangle>();
                _Diamonds = new List<Diamond>();
                _Triangles = new List<Triangle>();
                _Lines = new List<Line>();
                _CurvedLines = new List<CurvedLine>();
                _Arrows = new List<Arrow>();
                _DoubleArrows = new List<DoubleArrow>();
                _DoubleBraces = new List<DoubleBrace>();
                _DoubleBrackets = new List<DoubleBracket>();
                _Pentagons = new List<Pentagon>();
                _Hexagons = new List<Hexagon>();
                _Trapezoids = new List<Trapezoid>();
                _Pies = new List<Pie>();
                _Circles = new List<Circle>();
                _Images = new List<Image>();
                _Tables = new List<Table>();
                _SlideFacade.BackgroundColor = _BackgroundColor;
                _CommentFacade = new CommentFacade();
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Initializing slide");
                throw new Common.OpenizeException(errorMessage, ex);
            }

        }
        /// <summary>
        /// Contructor which accepts bool value 
        /// </summary>
        /// <param name="isNewSlide"></param>
        public Slide (bool isNewSlide)
        {
            if (isNewSlide)
                _SlideFacade = new SlideFacade(true);
            else
                _SlideFacade = new SlideFacade(false);


            _RelationshipId = _SlideFacade.RelationshipId;
            _TextShapes = new List<TextShape>();
            _Rectangles = new List<Rectangle>();
            _Diamonds = new List<Diamond>();
            _Triangles = new List<Triangle>();
            _Circles = new List<Circle>();
            _Lines = new List<Line>();
            _CurvedLines = new List<CurvedLine>();
            _Arrows = new List<Arrow>();
            _DoubleArrows = new List<DoubleArrow>();
            _DoubleBraces = new List<DoubleBrace>();
            _DoubleBrackets = new List<DoubleBracket>();
            _Pentagons = new List<Pentagon>();
            _Hexagons = new List<Hexagon>();
            _Trapezoids = new List<Trapezoid>();
            _Pies = new List<Pie>();
            _Images = new List<Image>();
            _Tables = new List<Table>();
            _CommentFacade = new CommentFacade();
        }
        /// <summary>
        /// Method to add a text shape in a slide.
        /// </summary>
        /// <param name="textShape">An object of TextShape class.</param>
        public void AddTextShapes (TextShape textShape)
        {
            try
            {
                if (textShape.TextList == null)
                {
                    textShape.Facade = _SlideFacade.AddTextShape(textShape.Text, textShape.FontSize, TextAlignment.Center,
                        Utility.PixelsToEmu(textShape.X), Utility.PixelsToEmu(textShape.Y)
                        , Utility.PixelsToEmu(textShape.Width), Utility.PixelsToEmu(textShape.Height), textShape.FontFamily,
                        textShape.TextColor, textShape.BackgroundColor);
                }
                else
                {
                    textShape.Facade = _SlideFacade.AddTextListShape(textShape.TextList.ListItems, textShape.TextList.Facade, textShape.FontSize, TextAlignment.Center,
                        Utility.PixelsToEmu(textShape.X), Utility.PixelsToEmu(textShape.Y)
                        , Utility.PixelsToEmu(textShape.Width), Utility.PixelsToEmu(textShape.Height), textShape.FontFamily,
                        textShape.TextColor, textShape.BackgroundColor);
                }
                textShape.ShapeIndex = _TextShapes.Count + 1;
                _TextShapes.Add(textShape);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding text shape");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }
        /// <summary>
        /// Method to draw a rectangular shape
        /// </summary>
        /// <param name="rect"></param>
        /// <exception cref="Common.OpenizeException"></exception>
        public void DrawRectangle(Rectangle rect)
        {
            try
            {
                var facade = new RectangleShapeFacade();
                facade.BackgroundColor = rect.BackgroundColor;
                rect.Facade = _SlideFacade.DrawRectangle(Utility.PixelsToEmu(rect.X), Utility.PixelsToEmu(rect.Y),
                    Utility.PixelsToEmu(rect.Width), Utility.PixelsToEmu(rect.Height), facade.BackgroundColor, facade);
                rect.ShapeIndex = _Rectangles.Count + 1;
                _Rectangles.Add(rect);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding rectangular shape");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Method to draw a triangular shape
        /// </summary>
        /// <param name="triangle"></param>
        /// <exception cref="Common.OpenizeException"></exception>
        public void DrawTriangle(Triangle triangle)
        {
            try
            {
                var facade = new TriangleShapeFacade();
                facade.BackgroundColor = triangle.BackgroundColor;
                triangle.Facade = _SlideFacade.DrawTriangle(Utility.PixelsToEmu(triangle.X), Utility.PixelsToEmu(triangle.Y),
                    Utility.PixelsToEmu(triangle.Width), Utility.PixelsToEmu(triangle.Height), facade.BackgroundColor, facade);
                triangle.ShapeIndex = _Triangles.Count + 1;
                _Triangles.Add(triangle);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding triangular shape");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Method to draw a diamond shape
        /// </summary>
        /// <param name="diamond"></param>
        /// <exception cref="Common.OpenizeException"></exception>
        public void DrawDiamond(Diamond diamond)
        {
            try
            {
                var facade = new DiamondShapeFacade();
                facade.BackgroundColor = diamond.BackgroundColor;
                diamond.Facade = _SlideFacade.DrawDiamond(Utility.PixelsToEmu(diamond.X), Utility.PixelsToEmu(diamond.Y),
                    Utility.PixelsToEmu(diamond.Width), Utility.PixelsToEmu(diamond.Height), facade.BackgroundColor, facade);
                diamond.ShapeIndex = _Diamonds.Count + 1;
                _Diamonds.Add(diamond);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding diamond shape");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Method to draw a line shape
        /// </summary>
        /// <param name="line"></param>
        /// <exception cref="Common.OpenizeException"></exception>
        public void DrawLine(Line line)
        {
            try
            {
                var facade = new LineFacade();
                line.Facade = _SlideFacade.DrawLine(Utility.PixelsToEmu(line.X), Utility.PixelsToEmu(line.Y),
                    Utility.PixelsToEmu(line.Width), Utility.PixelsToEmu(line.Height), facade);
                line.ShapeIndex = _Lines.Count + 1;
                _Lines.Add(line);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding line shape");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Method to draw a curved line
        /// </summary>
        /// <param name="curvedLine"></param>
        /// <exception cref="Common.OpenizeException"></exception>
        public void DrawCurvedLine(CurvedLine curvedLine)
        {
            try
            {
                var facade = new CurvedLineFacade();
                curvedLine.Facade = _SlideFacade.DrawCurvedLine(Utility.PixelsToEmu(curvedLine.X), Utility.PixelsToEmu(curvedLine.Y),
                    Utility.PixelsToEmu(curvedLine.Width), Utility.PixelsToEmu(curvedLine.Height), facade);
                curvedLine.ShapeIndex = _CurvedLines.Count + 1;
                _CurvedLines.Add(curvedLine);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding curved line shape");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Method to draw an arrow shape
        /// </summary>
        /// <param name="arrow"></param>
        /// <exception cref="Common.OpenizeException"></exception>
        public void DrawArrow(Arrow arrow)
        {
            try
            {
                var facade = new ArrowFacade();
                facade.Animation = arrow.Animation;
                arrow.Facade = _SlideFacade.DrawArrow(Utility.PixelsToEmu(arrow.X), Utility.PixelsToEmu(arrow.Y),
                    Utility.PixelsToEmu(arrow.Width), Utility.PixelsToEmu(arrow.Height), facade);
                arrow.ShapeIndex = _Arrows.Count + 1;
                _Arrows.Add(arrow);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding arrow shape");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Method to draw a double arrow shape
        /// </summary>
        /// <param name="doubleArrow"></param>
        /// <exception cref="Common.OpenizeException"></exception>
        public void DrawDoubleArrow(DoubleArrow doubleArrow)
        {
            try
            {
                var facade = new DoubleArrowFacade();
                facade.Animation = doubleArrow.Animation;
                doubleArrow.Facade = _SlideFacade.DrawDoubleArrow(Utility.PixelsToEmu(doubleArrow.X), Utility.PixelsToEmu(doubleArrow.Y),
                    Utility.PixelsToEmu(doubleArrow.Width), Utility.PixelsToEmu(doubleArrow.Height), facade);
                doubleArrow.ShapeIndex = _DoubleArrows.Count + 1;
                _DoubleArrows.Add(doubleArrow);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding double arrow shape");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Method to draw a double brace shape
        /// </summary>
        /// <param name="doubleBrace"></param>
        /// <exception cref="Common.OpenizeException"></exception>
        public void DrawDoubleBrace(DoubleBrace doubleBrace)
        {
            try
            {
                var facade = new DoubleBraceFacade();
                facade.Animation = doubleBrace.Animation;
                doubleBrace.Facade = _SlideFacade.DrawDoubleBrace(Utility.PixelsToEmu(doubleBrace.X), Utility.PixelsToEmu(doubleBrace.Y),
                    Utility.PixelsToEmu(doubleBrace.Width), Utility.PixelsToEmu(doubleBrace.Height), facade);
                doubleBrace.ShapeIndex = _DoubleBraces.Count + 1;
                _DoubleBraces.Add(doubleBrace);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding double brace shape");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Method to draw a double bracket shape
        /// </summary>
        /// <param name="doubleBracket"></param>
        /// <exception cref="Common.OpenizeException"></exception>
        public void DrawDoubleBracket(DoubleBracket doubleBracket)
        {
            try
            {
                var facade = new DoubleBracketFacade();
                facade.Animation= doubleBracket.Animation;
                doubleBracket.Facade = _SlideFacade.DrawDoubleBracket(Utility.PixelsToEmu(doubleBracket.X), Utility.PixelsToEmu(doubleBracket.Y),
                    Utility.PixelsToEmu(doubleBracket.Width), Utility.PixelsToEmu(doubleBracket.Height), facade);
                doubleBracket.ShapeIndex = _DoubleBrackets.Count + 1;
                _DoubleBrackets.Add(doubleBracket);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding double bracket shape");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Method to draw a pentagon shape
        /// </summary>
        /// <param name="pentagon"></param>
        /// <exception cref="Common.OpenizeException"></exception>
        public void DrawPentagon(Pentagon pentagon)
        {
            try
            {
                var facade = new PentagonFacade();
                facade.Animation= pentagon.Animation;
                pentagon.Facade = _SlideFacade.DrawPentagon(Utility.PixelsToEmu(pentagon.X), Utility.PixelsToEmu(pentagon.Y),
                    Utility.PixelsToEmu(pentagon.Width), Utility.PixelsToEmu(pentagon.Height), facade);
                pentagon.ShapeIndex = _Pentagons.Count + 1;
                _Pentagons.Add(pentagon);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding pentagon shape");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Method to draw a hexagon shape
        /// </summary>
        /// <param name="hexagon"></param>
        /// <exception cref="Common.OpenizeException"></exception>
        public void DrawHexagon(Hexagon hexagon)
        {
            try
            {
                var facade = new HexagonFacade();
                facade.Animation = hexagon.Animation;
                hexagon.Facade = _SlideFacade.DrawHexagon(Utility.PixelsToEmu(hexagon.X), Utility.PixelsToEmu(hexagon.Y),
                    Utility.PixelsToEmu(hexagon.Width), Utility.PixelsToEmu(hexagon.Height), facade);
                hexagon.ShapeIndex = _Hexagons.Count + 1;
                _Hexagons.Add(hexagon);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding hexagon shape");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Method to draw a trapezoid shape
        /// </summary>
        /// <param name="trapezoid"></param>
        /// <exception cref="Common.OpenizeException"></exception>
        public void DrawTrapezoid(Trapezoid trapezoid)
        {
            try
            {
                var facade = new TrapezoidFacade();
                facade.Animation= trapezoid.Animation;
                trapezoid.Facade = _SlideFacade.DrawTrapezoid(Utility.PixelsToEmu(trapezoid.X), Utility.PixelsToEmu(trapezoid.Y),
                    Utility.PixelsToEmu(trapezoid.Width), Utility.PixelsToEmu(trapezoid.Height), facade);
                trapezoid.ShapeIndex = _Trapezoids.Count + 1;
                _Trapezoids.Add(trapezoid);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding trapezoid shape");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        public void DrawPie(Pie pie)
        {
            try
            {
                var Facade= new PieFacade();
                Facade.Animation= pie.Animation;
               pie.Facade = _SlideFacade.DrawPie(Utility.PixelsToEmu(pie.X), Utility.PixelsToEmu(pie.Y),
                   Utility.PixelsToEmu(pie.Width), Utility.PixelsToEmu(pie.Height), Facade);
                pie.ShapeIndex = _Pies.Count + 1;
                Pies.Add(pie);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding text shape");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }
        /// <summary>
        /// Method to draw a circular shape
        /// </summary>
        /// <param name="circle"></param>
        /// <exception cref="Common.OpenizeException"></exception>
        public void DrawCircle(Circle circle)
        {
            try
            {
                var facade = new CircleShapeFacade();
                facade.Animation= circle.Animation;
                facade.BackgroundColor = circle.BackgroundColor;
                circle.Facade = _SlideFacade.DrawCircle(Utility.PixelsToEmu(circle.X), Utility.PixelsToEmu(circle.Y),
                    Utility.PixelsToEmu(circle.Width), Utility.PixelsToEmu(circle.Height), facade.BackgroundColor, facade);
                circle.ShapeIndex = _Circles.Count + 1;
                _Circles.Add(circle);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding circular shape");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Method to add/update note to a slide
        /// </summary>
        /// <param name="noteText">Text you want to add as note</param>
        public void AddNote(String noteText)
        {
            _SlideFacade.AddNote(noteText);
        }
        /// <summary>
        /// Method to remove Notes of a slide
        /// </summary>
        public void RemoveNote()
        {
            _SlideFacade.RemoveNote();
        }
        public void AddTextShapes (TextShape textShape, List<TextSegment> textSegments)
        {
            try
            {
                List<TextSegmentFacade> textSegmentFacades = new List<TextSegmentFacade>();
                foreach (var ts in textSegments)
                {
                    textSegmentFacades.Add(ts.Facade);
                }
                textShape.Facade = _SlideFacade.AddTextShape(textSegmentFacades, TextAlignment.Center,
                    Utility.PixelsToEmu(textShape.X), Utility.PixelsToEmu(textShape.Y)
                    , Utility.PixelsToEmu(textShape.Width), Utility.PixelsToEmu(textShape.Height), textShape.BackgroundColor);
                textShape.ShapeIndex = _TextShapes.Count + 1;
                _TextShapes.Add(textShape);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding text shape");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }
        /// <summary>
        /// Method to add images to a slide. 
        /// </summary>
        /// <param name="image">An object of Image class</param>
        public void AddImage (Image image)
        {
            try
            {
                image.Facade = new ImageFacade();
                image.Facade.X = Utility.PixelsToEmu(image.X);
                image.Facade.Y = Utility.PixelsToEmu(image.Y);
                image.Facade.Width = Utility.PixelsToEmu(image.Width);
                image.Facade.Height = Utility.PixelsToEmu(image.Height);
                image.Facade.createImage(image.ImagePath, _SlideFacade.SlidePart);
                _SlideFacade.AddImage(image.Facade);
                image.ImageIndex = _Images.Count + 1;
                _Images.Add(image);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding image");
                throw new Common.OpenizeException(errorMessage, ex);
            }

        }
        /// <summary>
        /// Method to add comments to a slide. 
        /// </summary>
        /// <param name="comment">An object of Comment class</param>
        public void AddComment(Comment comment)
        {
            try
            {
                comment.Facade = _CommentFacade;
                if (_CommentIndex == 0)
                {                    
                    comment.Facade.SetAssociatedSlidePart(_SlideFacade.SlidePart, _SlidePresentation.Facade.CommentAuthorPart);
                }
                UInt32Value id = new UInt32Value { Value = (uint)comment.AuthorId };
                comment.Facade.GenerateComment(id,comment.Text, comment.InsertedAt, comment.X, comment.Y);
                _CommentIndex += 1;
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding Comment");
                throw new Common.OpenizeException(errorMessage, ex);
            }

        }
        /// <summary>
        /// Method to get the list of comments.
        /// </summary>
        /// <returns></returns>
        public List<Comment> GetComments()
        {
            var comments= new List<Comment>();
            var facadeComments= _SlideFacade.GetComments();
            foreach (var facadeComment in facadeComments)
            {
                Comment comment = new Comment();
                comment.Text = facadeComment["Text"];
                comment.AuthorId = Convert.ToInt32(facadeComment["AuthorId"]);
                comment.CommentIndex= Convert.ToInt32(facadeComment["Index"]);
                comment.InsertedAt = Convert.ToDateTime(facadeComment["DateTime"]);
                comment.X = Convert.ToInt64(facadeComment["XPos"]);
                comment.Y = Convert.ToInt64(facadeComment["YPos"]);
                comment.Facade = new CommentFacade();
                comment.Facade.CommentPart = _SlideFacade.CommentPart;

                comments.Add(comment);
            }
            return comments;
        }
        /// <summary>
        /// Method to add table to a slide. 
        /// </summary>
        /// <param name="table">An object of Table class</param>
        public void AddTable (Table table)
        {
            try
            {
                table.Facade = new TableFacade();
                table.Facade.X = Utility.PixelsToEmu(table.X);
                table.Facade.Y = Utility.PixelsToEmu(table.Y);
                table.Facade.Width = Utility.PixelsToEmu(table.Width);
                table.Facade.Height = Utility.PixelsToEmu(table.Height);
                if (table.Theme == null)
                {
                    table.Theme = Table.TableStyle.LightStyle1;
                }
                
                table.Facade.GenerateTable(_SlideFacade.SlidePart, table.GetDataTable());
                table.TableIndex = _Tables.Count + 1;
                _Tables.Add(table);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Adding table");
                throw new Common.OpenizeException(errorMessage, ex);
            }

        }        

        /// <summary>
        /// Get text shapes by searching a text term.
        /// </summary>
        /// <param name="text">Search term as string</param>
        /// <returns></returns>
        public List<TextShape> GetTextShapesByText (String text)
        {
            try
            {
                List<TextShape> shapes = TextShapes.Where(shape => shape.Text.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
                return shapes;
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Getting Shapes");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }
        
        /// <summary>
        /// Method to update a slide properties e.g. background color.
        /// </summary>
        public void Update ()
        {
            _SlideFacade.BackgroundColor = _BackgroundColor;
            _SlideFacade.Update();
        }

    }
}
