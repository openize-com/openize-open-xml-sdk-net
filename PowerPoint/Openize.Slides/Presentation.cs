using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Openize.Slides.Facade;
using System;
using System.Collections.Generic;

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
        private string _FilePath = null;

        /// <summary>
        /// Gets the facade for handling presentation document operations.
        /// </summary>
        public PresentationDocumentFacade Facade { get => doc; }

        /// <summary>
        /// Gets or sets the slide width of the presentation.
        /// </summary>
        public int SlideWidth { get => _SlideWidth; set => _SlideWidth = value; }

        /// <summary>
        /// Gets or sets the slide height of the presentation.
        /// </summary>
        public int SlideHeight { get => _SlideHeight; set => _SlideHeight = value; }

        /// <summary>
        /// Gets or sets the file path of the presentation.
        /// </summary>
        public string FilePath { get => _FilePath; set => _FilePath = value; }

        /// <summary>
        /// Initializes a new presentation document with the specified file path.
        /// </summary>
        /// <param name="FilePath">The file path for the presentation.</param>
        private Presentation(String FilePath)
        {
            try
            {
                _FilePath = FilePath;
                _Slides = new List<Slide>();
                _CommentAuthors = new List<CommentAuthor>();
                doc = PresentationDocumentFacade.Create(FilePath);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Initializing presentation");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Initializes an empty presentation.
        /// </summary>
        private Presentation()
        {
            try
            {
                _Slides = new List<Slide>();
                _CommentAuthors = new List<CommentAuthor>();
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Initializing empty presentation");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Creates a new presentation instance with the specified file path.
        /// </summary>
        /// <param name="FilePath">The file path for the new presentation.</param>
        public static Presentation Create(String FilePath)
        {
            return new Presentation(FilePath);
        }

        /// <summary>
        /// Opens an existing presentation file.
        /// </summary>
        /// <param name="FilePath">The file path of the presentation to open.</param>
        public static Presentation Open(String FilePath)
        {
            try
            {
                doc = PresentationDocumentFacade.Open(FilePath);
                return new Presentation();
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Opening presentation");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Appends a slide to the presentation.
        /// </summary>
        /// <param name="slide">The slide object to append.</param>
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
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Appending slide");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Copies a slide from another presentation.
        /// </summary>
        /// <param name="slide">The slide object to copy.</param>
        public void CopySlide(Slide slide)
        {
            try
            {
                doc.CopySlide(slide.SlideFacade);
                _Slides.Add(slide);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Copying slide");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Retrieves all slides from the presentation.
        /// </summary>
        public List<Slide> GetSlides()
        {
            try
            {
                if (!doc.IsNewPresentation)
                {
                    foreach (var slidepart in doc.PresentationSlideParts)
                    {
                        var slide = new Slide(false);
                        slide.SlideFacade = new SlideFacade(false);
                        _Slides.Add(slide);
                    }
                }
                return _Slides;
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Getting slides");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Saves the presentation document.
        /// </summary>
        public void Save()
        {
            try
            {
                doc.Close(FilePath);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Saving presentation");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }

        /// <summary>
        /// Closes the presentation document.
        /// </summary>
        public void Close()
        {
            try
            {
                doc.Close(FilePath);
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Closing presentation");
                throw new Common.OpenizeException(errorMessage, ex);
            }
        }
    }
}