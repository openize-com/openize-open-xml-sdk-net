using DocumentFormat.OpenXml.Presentation;
using Openize.Slides.Facade;
using System;
using System.Collections.Generic;

namespace Openize.Slides
{
    /// <summary>
    /// Represents the presentation document.
    /// </summary>
    public class Presentation
    {
        private static String _FileName = "MyPresentation.pptx";
        private static String _DirectoryPath = "D:\\AsposeSampleResults\\";
        private static PresentationDocumentFacade doc = null;
        private List<Slide> _Slides;
        private List<CommentAuthor> _CommentAuthors;

        public PresentationDocumentFacade Facade { get => doc; }



        /// <summary>
        /// Initializes the presentation object.
        /// </summary>
        /// <param name="FilePath">Presentation path as string</param>
        private Presentation (String FilePath)
        {
            _Slides = new List<Slide>();
            _CommentAuthors = new List<CommentAuthor>();
            doc = PresentationDocumentFacade.Create(FilePath);

        }
        /// <summary>
        /// Default constructor to initialize presentation object.
        /// </summary>
        private Presentation ()
        {
            _Slides = new List<Slide>();
            _CommentAuthors = new List<CommentAuthor>();

        }
        /// <summary>
        /// Static method to instantiate a new object of Presentation class.
        /// </summary>
        /// <param name="FilePath">Presentation path as string</param>
        /// <returns>An instance of Presentation object</returns>
        /// Use this method to create a new, blank presentation that you can populate with content.
        /// To work with an existing document, consider using the <see cref="Open(string)"/> method.
        /// <example>
        /// <code>
        /// Presentation presentation = Presentation.Create("D:\\AsposeSampleResults\\test2.pptx");
        /// TextShape shape = new TextShape();
        /// shape.Text = "Title: Here is my first title From FF";
        /// TextShape shape2 = new TextShape();
        /// shape2.Text = "Body : Here is my first title From FF";
        ///  // First slide
        /// Slide slide = new Slide();
        /// slide.AddTextShapes(shape);
        /// slide.AddTextShapes(shape2);
        /// // 2nd slide
        /// Slide slide1 = new Slide();
        /// slide1.AddTextShapes(shape);
        /// slide1.AddTextShapes(shape2);
        /// // Adding slides
        /// presentation.AppendSlide(slide);
        /// presentation.AppendSlide(slide1);
        /// presentation.Save();
        /// </code>
        /// </example>
        public static Presentation Create (String FilePath)
        {
            return new Presentation(FilePath);
        }
        /// <summary>
        /// Static method to load an existing presentation.
        /// </summary>
        /// <param name="FilePath">Presentation path as string</param>
        /// <returns></returns>
        /// <example>
        /// <code>
        /// Presentation presentation = Presentation.Open("D:\\AsposeSampleData\\sample.pptx");
        /// TextShape shape1 = new TextShape();
        /// shape1.Text = "Title: Here is my first title From FF";
        /// TextShape shape2 = new TextShape();
        /// shape2.Text = "Body : Here is my first title From FF";
        ///  // New slide
        /// Slide slide = new Slide();
        /// slide.AddTextShapes(shape1);
        /// slide.AddTextShapes(shape2);       
        /// // Adding slide
        /// presentation.AppendSlide(slide);
        /// presentation.Save();
        /// </code>
        /// </example>
        public static Presentation Open (String FilePath)
        {
            doc = PresentationDocumentFacade.Open(FilePath);
            return new Presentation();
        }
        /// <summary>
        /// Create comment author using this method
        /// </summary>
        /// <param name="author"> Pass comment author object</param>
        public void CreateAuthor(CommentAuthor author)
        {
            doc.CreateAuthor(author.Id, author.ColorIndex, author.Name, author.InitialLetter);
            _CommentAuthors.Add(author);
        }
        /// <summary>
        /// Get the list of comment author
        /// </summary>
        /// <returns></returns>
        public List<CommentAuthor> GetCommentAuthors()
        {
            List<CommentAuthor > authorList = new List<CommentAuthor>();
            var FacadeAuthors = doc.GetCommentAuthors();
            foreach(var author in FacadeAuthors)
            {
                CommentAuthor commentAuthor = new CommentAuthor();
                commentAuthor.InitialLetter = author["Initials"];
                commentAuthor.ColorIndex = Convert.ToInt32(author["ColorIndex"]);
                commentAuthor.Name = author["Name"];
                commentAuthor.Id = Convert.ToInt32(author["Id"]);  
                authorList.Add(commentAuthor);
            }
            return authorList;
        }
        /// <summary>
        /// This method is responsible to append a slide.
        /// </summary>
        /// <param name="slide">An object of a slide</param>
        public void AppendSlide (Slide slide)
        {
            slide.SlideFacade.SetSlideBackground(slide.BackgroundColor);
            doc.AppendSlide(slide.SlideFacade);
            _Slides.Add(slide);

        }
        /// <summary>
        /// Method to get the list of all slides of a presentation
        /// </summary>
        /// <returns></returns>
        /// <example>
        /// <code>
        /// Presentation presentation = Presentation.Open("D:\\AsposeSampleData\\sample.pptx");
        /// var slides = presentation.GetSlides();
        /// var slide = slides[0];
        /// ...
        /// </code>
        /// </example>
        public List<Slide> GetSlides ()
        {
            if (!doc.IsNewPresentation)
            {
                foreach (var slidepart in doc.PresentationSlideParts)
                {
                    var slide = new Slide(false);
                    
                    SlideFacade slideFacade = new SlideFacade(false);
                    slideFacade.TextShapeFacades = TextShapeFacade.PopulateTextShapes(slidepart);
                    slideFacade.RectangleShapeFacades = RectangleShapeFacade.PopulateRectangleShapes(slidepart);
                    slideFacade.ImagesFacade = ImageFacade.PopulateImages(slidepart);
                    slideFacade.PresentationSlide = slidepart.Slide;
                    slideFacade.TableFacades= TableFacade.PopulateTables(slidepart);
                    slideFacade.SlidePart = slidepart;
                    slideFacade.CommentPart = slidepart.SlideCommentsPart;
                    slideFacade.NotesPart = slidepart.NotesSlidePart;
                    slideFacade.RelationshipId = doc.GetSlideRelationshipId(slidepart);
                    slide.TextShapes = TextShape.GetTextShapes(slideFacade.TextShapeFacades);
                    slide.Rectangles = Rectangle.GetRectangles(slideFacade.RectangleShapeFacades);
                    slide.Circles = Circle.GetCircles(slideFacade.CircleShapeFacades);
                    slide.Images = Image.GetImages(slideFacade.ImagesFacade);
                    slide.Tables = Table.GetTables(slideFacade.TableFacades);
                    slide.SlideFacade = slideFacade;
                    slide.SlidePresentation = this;
                    _Slides.Add(slide);
                }
            }
            return _Slides;

        }
        /// <summary>
        /// Extract and save images of a presentation into a director
        /// </summary>
        /// <param name="outputFolder">Folder path as string</param>
        public void ExtractAndSaveImages (string outputFolder)
        {
            doc.ExtractAndSaveImages(outputFolder);
        }
        /// <summary>
        /// Method to remove a slide at a specific index
        /// </summary>
        /// <param name="slideIndex">Index of a slide</param>  
        /// <example>
        /// <code>
        /// Presentation presentation = Presentation.Open("D:\\AsposeSampleData\\sample.pptx");
        /// var confirmation = presentation.RemoveSlide(0);
        /// Console.WriteLine(confirmation);
        /// presentation.Save();
        /// </code>
        /// </example>
        public String RemoveSlide (int slideIndex)
        {
            return doc.RemoveSlide(slideIndex);
        }
        /// <summary>
        /// Method to remove comment author.
        /// </summary>
        /// <param name="author"></param>
        public void RemoveCommentAuthor(CommentAuthor author)
        {
            doc.RemoveCommentAuthor(author.Id);
            _CommentAuthors.Remove(author);
        }
        /// <summary>
        /// Method to insert a slide at a specific index
        /// </summary>
        /// <param name="index">Index of a slide</param>
        /// <param name="slide">A slide object</param>
        public void InsertSlideAt (int index, Slide slide)
        {
            slide.SlideIndex = index;
            slide.SlideFacade.SlideIndex = index;
            doc.InsertSlide(index, slide.SlideFacade);
        }

        /// <summary>
        /// This method exports all existing notes of a PPT/PPTX to TXT file.
        /// </summary>
        /// <param name="filePath"> File path where to save TXT file</param>
        public void SaveAllNotesToTextFile(string filePath)
        {
            doc.SaveAllNotesToTextFile(filePath);
        }
        /// <summary>
        /// Method to save the new or changed presentation.
        /// </summary>
        public void Save ()
        {
            doc.Save();

        }
    }
}
