using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Openize.Slides.Facade;
using System;
using System.Collections.Generic;
using System.IO;

namespace Openize.Slides
{
    public class Presentation
    {
        private static String _FileName = "MyPresentation.pptx";
        private static String _DirectoryPath = "D:\\AsposeSampleResults\\";
        private static PresentationDocumentFacade doc = null;
        private List<Slide> _Slides = null;
        private List<CommentAuthor> _CommentAuthors = null;
        private int _SlideWidth = 960;
        private int _SlideHeight = 720;

        public PresentationDocumentFacade Facade => doc;
        /// <summary>
        /// Set slide width
        /// </summary>
        public int SlideWidth { get => _SlideWidth; set => _SlideWidth = value; }
        /// <summary>
        /// Set slide height
        /// </summary>
        public int SlideHeight { get => _SlideHeight; set => _SlideHeight = value; }

        private Presentation(String FilePath)
        {
            _Slides = new List<Slide>();
            _CommentAuthors = new List<CommentAuthor>();

            try
            {
                doc = PresentationDocumentFacade.Create(FilePath);
            }
            catch (Exception ex)
            {
                throw new Common.OpenizeException($"Failed to create presentation at path '{FilePath}'.", ex);
            }
        }

        private Presentation()
        {
            _Slides = new List<Slide>();
            _CommentAuthors = new List<CommentAuthor>();
        }

        /// <summary>
        /// Static method to instantiate a new object of Presentation class.
        /// </summary>
        /// <param name="FilePath">Presentation path as string</param>
        /// <returns>An instance of Presentation object</returns>
        public static Presentation Create(String FilePath)
        {
            try
            {
                return new Presentation(FilePath);
            }
            catch (Exception ex)
            {
                throw new Common.OpenizeException($"Error occurred while creating a new presentation: {FilePath}", ex);
            }
        }

        /// <summary>
        /// Static method to load an existing presentation.
        /// </summary>
        /// <param name="FilePath">Presentation path as string</param>
        /// <returns>Instance of Presentation object</returns>
        public static Presentation Open(String FilePath)
        {
            try
            {
                doc = PresentationDocumentFacade.Open(FilePath);
                return new Presentation();
            }
            catch (Exception ex)
            {
                throw new Common.OpenizeException($"Error occurred while opening the presentation: {FilePath}", ex);
            }
        }

        /// <summary>
        /// This method is responsible to append a slide.
        /// </summary>
        /// <param name="slide">An object of a slide</param>
        public void AppendSlide(Slide slide)
        {
            try
            {
                slide.SlideFacade.SetSlideBackground(slide.BackgroundColor);
                doc.SlideWidth = new Int32Value((int)Common.Utility.PixelsToEmu(SlideWidth));
                doc.SlideHeight = new Int32Value((int)Common.Utility.PixelsToEmu(SlideHeight));
                doc.AppendSlide(slide.SlideFacade);
                _Slides.Add(slide);
            }
            catch (Exception ex)
            {
                throw new Common.OpenizeException("Failed to append slide to presentation.", ex);
            }
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
        public List<Slide> GetSlides()
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
                    slideFacade.TableFacades = TableFacade.PopulateTables(slidepart);
                    slideFacade.SlidePart = slidepart;
                    slideFacade.CommentPart = slidepart.SlideCommentsPart;
                    slideFacade.NotesPart = slidepart.NotesSlidePart;
                    slideFacade.RelationshipId = doc.GetSlideRelationshipId(slidepart);
                    slide.SetTextShapesDirect(TextShape.GetTextShapes(slideFacade.TextShapeFacades));
                    slide.SetRectanglesDirect(Rectangle.GetRectangles(slideFacade.RectangleShapeFacades));
                    slide.SetCirclesDirect(Circle.GetCircles(slideFacade.CircleShapeFacades));
                    slide.SetImagesDirect(Image.GetImages(slideFacade.ImagesFacade));
                    slide.SetTablesDirect(Table.GetTables(slideFacade.TableFacades));
                    slide.SlideFacade = slideFacade;
                    slide.SlidePresentation = this;
                    _Slides.Add(slide);
                }
            }
            return _Slides;

        }

        /// <summary>
        /// Method to save the new or changed presentation.
        /// </summary>
        public void Save()
        {
            try
            {
                doc.Save();
            }
            catch (Exception ex)
            {
                throw new Common.OpenizeException("Failed to save the presentation.", ex);
            }
        }

        /// <summary>
        /// Method to close a presentation.
        /// </summary>
        public void close()
        {
            try
            {
                doc.Dispose();
            }
            catch (Exception ex)
            {
                throw new Common.OpenizeException("Failed to close the presentation.", ex);
            }
        }

        /// <summary>
        /// Extract and save images of a presentation into a directory.
        /// </summary>
        /// <param name="outputFolder">Folder path as string</param>
        public void ExtractAndSaveImages(string outputFolder)
        {
            try
            {
                doc.ExtractAndSaveImages(outputFolder);
            }
            catch (Exception ex)
            {
                throw new Common.OpenizeException($"Failed to extract and save images to folder '{outputFolder}'.", ex);
            }
        }

        /// <summary>
        /// Method to remove a slide at a specific index.
        /// </summary>
        /// <param name="slideIndex">Index of a slide</param>
        public String RemoveSlide(int slideIndex)
        {
            try
            {
                return doc.RemoveSlide(slideIndex);
            }
            catch (Exception ex)
            {
                throw new Common.OpenizeException($"Failed to remove slide at index {slideIndex}.", ex);
            }
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
            List<CommentAuthor> authorList = new List<CommentAuthor>();
            var FacadeAuthors = doc.GetCommentAuthors();
            foreach (var author in FacadeAuthors)
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
        public void InsertSlideAt(int index, Slide slide)
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
    }
}
