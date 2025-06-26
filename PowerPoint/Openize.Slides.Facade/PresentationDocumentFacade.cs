using System;
using System.Collections.Generic;
using System.Text;
using D = DocumentFormat.OpenXml.Drawing;
using System.IO;
using DocumentFormat.OpenXml.Presentation;
using PKG = DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using System.Linq;
using Openize.Slides.Common;
using System.Dynamic;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace Openize.Slides.Facade
{
    public class PresentationDocumentFacade : IDisposable
    {
        private static readonly Dictionary<string, PresentationDocumentFacade> _instances = new Dictionary<string, PresentationDocumentFacade>();
        private static PresentationDocumentFacade _lastInstance;
        private static string _FilePath = null;
        private static MemoryStream _MemoryStream = null;
        private PKG.PresentationDocument _PresentationDocument = null;
        private bool disposedValue;
        private PKG.PresentationPart _PresentationPart = null;
        private List<PKG.SlideLayoutPart> _PresentationSlideLayoutParts = null;
        private List<SlidePart> _PresentationSlideParts = null;
        private SlideMasterPart _PresentationSlideMasterPart = null;
        private PKG.ThemePart _PresentationThemePart = null;
        private SlideIdList _SlideIdList = null;
        private bool isNewPresentation = false;
        private List<SlideFacade> _SlideFacades = null;
        private CommentAuthorsPart _CommentAuthorPart;
        private Int32Value _slideWidth;
        private Int32Value _slideHeight;




        public List<SlideLayoutPart> PresentationSlideLayoutParts { get => _PresentationSlideLayoutParts; set => _PresentationSlideLayoutParts = value; }
        public List<SlidePart> PresentationSlideParts { get => _PresentationSlideParts; set => _PresentationSlideParts = value; }
        public SlideMasterPart PresentationSlideMasterPart { get => _PresentationSlideMasterPart; set => _PresentationSlideMasterPart = value; }
        public ThemePart PresentationThemePart { get => _PresentationThemePart; set => _PresentationThemePart = value; }
        public SlideIdList SlideIdList { get => _SlideIdList; set => _SlideIdList = value; }
        public List<SlideFacade> SlideFacades { get => _SlideFacades; set => _SlideFacades = value; }
        public bool IsNewPresentation { get => isNewPresentation; set => isNewPresentation = value; }
        public CommentAuthorsPart CommentAuthorPart { get => _CommentAuthorPart; set => _CommentAuthorPart = value; }
        public Int32Value SlideWidth { get => _slideWidth; set => _slideWidth = value; }
        public Int32Value SlideHeight { get => _slideHeight; set => _slideHeight = value; }
        public static string FilePath { get => _FilePath; set => _FilePath = value; }

        public PKG.PresentationPart GetPresentationPart ()
        {
            return _PresentationPart;
        }
        private PresentationDocumentFacade (String FilePath, bool isNewFile)
        {

            try
            {
                if (isNewFile)
                {
                    _FilePath = FilePath;
                    IsNewPresentation = isNewFile;
                    SlideMasterIdList slideMasterIdList = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });

                    _PresentationDocument = PKG.PresentationDocument.Create(FilePath, PresentationDocumentType.Presentation);
                    _PresentationPart = _PresentationDocument.AddPresentationPart();
                    _PresentationPart.Presentation = new P.Presentation();
                   
                    _SlideIdList = new SlideIdList();
                    _PresentationSlideParts = new List<SlidePart>();
                    _PresentationSlideLayoutParts = new List<SlideLayoutPart>();
                    _PresentationPart.Presentation.Append(slideMasterIdList, _SlideIdList);
                    CreateCommentAuthorPart();
                }
                else
                {
                    _FilePath = FilePath;
                    _PresentationDocument = PKG.PresentationDocument.Open(FilePath, true);
                    _PresentationPart = _PresentationDocument.PresentationPart;
                    _PresentationSlideParts = GetSlideParts(_PresentationPart);
                    _CommentAuthorPart = _PresentationPart.CommentAuthorsPart;
                    _PresentationSlideLayoutParts = GetSlideLayoutParts(_PresentationSlideParts);
                    _SlideIdList = _PresentationPart.Presentation.SlideIdList;
                    _PresentationSlideMasterPart = _PresentationPart.SlideMasterParts.FirstOrDefault();


                    Utility.NextIndex = GetHighestNumericPart(_PresentationPart);

                }
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Loading Document");
                throw new Common.OpenizeException(errorMessage, ex);
            }


        }

        private List<SlideLayoutPart> GetSlideLayoutParts (List<SlidePart> presentationSlideParts)
        {
            var slideLayoutParts = new List<SlideLayoutPart>();

            foreach (var slidePart in presentationSlideParts)
            {
                slideLayoutParts.Add(slidePart.SlideLayoutPart);
                Utility.SlideNextIndex += 1;
            }

            return slideLayoutParts;
        }
        private List<SlidePart> GetSlideParts(PresentationPart _presentationPart)
        {
            List<SlidePart> SlideParts = new List<SlidePart>();
            foreach (SlideId slideId in _presentationPart.Presentation.SlideIdList)
            {
                // Get the relationship ID of the slide
                string relId = slideId.RelationshipId;

                // Get the slide part using the relationship ID
                SlidePart slidePart = (SlidePart)_presentationPart.GetPartById(relId);

                // Now you can work with the slidePart object
                // For example, you can add it to a List<SlidePart>

                SlideParts.Add(slidePart);
            }
            return SlideParts;
        }

        public IEnumerable<Dictionary<string, string>> GetCommentAuthors()
        {
            List<Dictionary<string, string>> commentAuthors = new List<Dictionary<string, string>>();

            
                // Get the list of comment authors
                CommentAuthorsPart commentAuthorsPart = _PresentationPart.CommentAuthorsPart;

                if (commentAuthorsPart != null)
                {
                    var commentAuthorList = commentAuthorsPart.CommentAuthorList;
                    // Extract comment authors
                    foreach (var author in commentAuthorList.Elements<P.CommentAuthor>())
                    {
                        Dictionary<string, string> authorProperties = new Dictionary<string, string>
                    {
                        { "Id", author.Id },
                        { "Name", author.Name },
                        { "Initials", author.Initials },
                        { "LastIndex", author.LastIndex },
                        { "ColorIndex", author.ColorIndex }
                    };

                        commentAuthors.Add(authorProperties);
                    }
                }
            

            return commentAuthors;
        }
        public void RemoveCommentAuthor(int id)
        {
            
            var commentAuthorToRemove = _CommentAuthorPart.CommentAuthorList.Descendants<P.CommentAuthor>()
                .FirstOrDefault(author => author.Id == id);
            if (commentAuthorToRemove != null)
            {
                commentAuthorToRemove.Remove();
            }
        }

        public static PresentationDocumentFacade Create (String FilePath)
        {

            if (!_instances.ContainsKey(FilePath))
            {
                _instances[FilePath] = new PresentationDocumentFacade(FilePath, true);
            }
            _lastInstance = _instances[FilePath];
            return _instances[FilePath];
        }

        public static PresentationDocumentFacade Create(String FilePath,int SlideWidth, int SlideHeight)
        {

            if (!_instances.ContainsKey(FilePath))
            {
                _instances[FilePath] = new PresentationDocumentFacade(FilePath, true);
            }
            _lastInstance = _instances[FilePath];
            return _instances[FilePath];
        }
        public static PresentationDocumentFacade Open (string FilePath)
        {
            if (!_instances.ContainsKey(FilePath))
            {
                _instances[FilePath] = new PresentationDocumentFacade(FilePath, false);
            }
            _lastInstance = _instances[FilePath];
            return _instances[FilePath];
        }

        public static PresentationDocumentFacade getInstance(string FilePath = null)
        {
            return FilePath != null ? _instances.GetValueOrDefault(FilePath) : _lastInstance;
        }

        private void CreatePresentationParts ()
        {
            // Default values in EMUs
            const int defaultWidth = 9144000; // 10 inches
            const int defaultHeight = 6858000; // 7.5 inches

            // Use the class-level values if set, otherwise use defaults
            Int32Value slideWidth = _slideWidth ?? defaultWidth;
            Int32Value slideHeight = _slideHeight ?? defaultHeight;

            //SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
            SlideSize slideSize1 = new SlideSize() { Cx = slideWidth, Cy = slideHeight, Type = SlideSizeValues.Custom };
            NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            _PresentationPart.Presentation.Append(slideSize1, notesSize1, defaultTextStyle1);



            CreateSlideLayoutPart();
            CreateSlideMasterPart();            
            CreateThemePart("rId5");

            _PresentationSlideMasterPart.AddPart(_PresentationSlideLayoutParts[0], "rId1");
            _PresentationPart.AddPart(_PresentationSlideMasterPart, "rId1");
            try
            {
                Utility.NextIndex += 1;
                var relId = Utility.GetUniqueRelationshipId();

                _PresentationPart.AddPart(_PresentationThemePart, "rId5");
            }
            catch (Exception ex)
            {
                Utility.NextIndex += 1;
                var relId = Utility.GetUniqueRelationshipId();
                _PresentationPart.AddPart(_PresentationThemePart, relId);

            }
        }

         private void CreateCommentAuthorPart()
        {
            CommentAuthorsPart commentAuthorsPart1 = _PresentationPart.AddNewPart<CommentAuthorsPart>("rId3");
            CommentAuthorList commentAuthorList1 = new CommentAuthorList();
            commentAuthorList1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            commentAuthorList1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            commentAuthorList1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
                    

            commentAuthorsPart1.CommentAuthorList = commentAuthorList1;

            _CommentAuthorPart = commentAuthorsPart1;
        }
        public void CreateAuthor(int id, int colorIndex, String name, string initialLetter)
        {

            UInt32Value authId = Convert.ToUInt32(id);
            UInt32Value color_index = new UInt32Value { Value = (uint)colorIndex };
            P.CommentAuthor commentAuthor1 = new P.CommentAuthor() { Id = authId, Name = name, Initials = initialLetter, LastIndex = authId, ColorIndex = color_index };

            CommentAuthorExtensionList commentAuthorExtensionList1 = new CommentAuthorExtensionList();

            CommentAuthorExtension commentAuthorExtension1 = new CommentAuthorExtension() { Uri = "{19B8F6BF-5375-455C-9EA6-DF929625EA0E}" };

            P15.PresenceInfo presenceInfo1 = new P15.PresenceInfo() { UserId = name, ProviderId = "None" };
            presenceInfo1.AddNamespaceDeclaration("p15", "http://schemas.microsoft.com/office/powerpoint/2012/main");

            commentAuthorExtension1.Append(presenceInfo1);

            commentAuthorExtensionList1.Append(commentAuthorExtension1);

            commentAuthor1.Append(commentAuthorExtensionList1);

            _CommentAuthorPart.CommentAuthorList.Append(commentAuthor1);
        }

        public void CreateSlideLayoutPart ()
        {

            int index = 1;
            foreach (var slidePart in _PresentationSlideParts)
            {
                if (slidePart.SlideLayoutPart == null)
                {
                    if (index == 1)
                    {


                        var _PresentationSlideLayoutPart = slidePart.AddNewPart<PKG.SlideLayoutPart>("rId1");

                        SlideLayout slideLayout = new SlideLayoutFacade().PresentationSlideLayout;
                        _PresentationSlideLayoutPart.SlideLayout = slideLayout;
                        _PresentationSlideLayoutParts.Add(_PresentationSlideLayoutPart);
                    }
                    else
                    {
                        slidePart.AddPart(_PresentationSlideLayoutParts[0]);
                    }
                    index++;
                }
            }
            //return slideLayoutPart1;
        }



        private void CreateSlideMasterPart ()
        {
            _PresentationSlideMasterPart = _PresentationSlideLayoutParts[0].AddNewPart<PKG.SlideMasterPart>("rId1");

            _PresentationSlideMasterPart.SlideMaster = new SlideMasterFacade().PresentationSlideMaster;

            //return slideMasterPart1;
        }

        private void CreateThemePart (string relId)
        {
            _PresentationThemePart = _PresentationSlideMasterPart.AddNewPart<PKG.ThemePart>(relId);

            _PresentationThemePart.Theme = new ThemeFacade().PresentationTheme;
            //return themePart1;

        }

        private int GetHighestNumericPart (PresentationPart presentationPart)
        {
            List<int> numericParts = new List<int>();

            if (presentationPart != null)
            {
                var slideIds = presentationPart.Presentation.SlideIdList;

                foreach (var slideId in slideIds.Elements<SlideId>())
                {
                    int slideNumericPart;
                    if (TryExtractNumericPart(slideId.RelationshipId, out slideNumericPart))
                    {
                        numericParts.Add(slideNumericPart);
                    }

                    SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;

                    if (slidePart != null)
                    {
                        // Add slide layout numeric part
                        int layoutNumericPart;
                        SlideLayoutPart layoutPart = slidePart.SlideLayoutPart;
                        if (layoutPart != null && TryExtractNumericPart(slidePart.GetIdOfPart(layoutPart), out layoutNumericPart))
                        {
                            numericParts.Add(layoutNumericPart);
                        }

                        // Add shape numeric parts
                        var shapeTree = slidePart.Slide.Descendants<ShapeTree>().FirstOrDefault();

                        if (shapeTree != null)
                        {
                            foreach (var shape in shapeTree.Descendants<D.NonVisualDrawingProperties>())
                            {
                                int shapeNumericPart;
                                string referenceId = shape.Id;

                                if (!string.IsNullOrEmpty(referenceId) && TryExtractNumericPart(referenceId, out shapeNumericPart))
                                {
                                    numericParts.Add(shapeNumericPart);
                                }
                            }
                        }
                    }
                }

                // Add slide master numeric part
                int slideMasterNumericPart;
                SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.FirstOrDefault();
                if (slideMasterPart != null && TryExtractNumericPart(presentationPart.GetIdOfPart(slideMasterPart), out slideMasterNumericPart))
                {
                    numericParts.Add(slideMasterNumericPart);
                }

                // Add theme numeric part
                int themeNumericPart;
                ThemePart themePart = presentationPart.ThemePart;
                if (themePart != null && TryExtractNumericPart(presentationPart.GetIdOfPart(themePart), out themeNumericPart))
                {
                    numericParts.Add(themeNumericPart);
                }

                // Add TableStylesPart numeric part
                int tableStylesNumericPart;
                TableStylesPart tableStylesPart = presentationPart.TableStylesPart;
                if (tableStylesPart != null && TryExtractNumericPart(presentationPart.GetIdOfPart(tableStylesPart), out tableStylesNumericPart))
                {
                    numericParts.Add(tableStylesNumericPart);
                }
            }

            return numericParts.Max();
        }
        private bool TryExtractNumericPart (string referenceId, out int numericPart)
        {
            numericPart = 0;

            if (referenceId == null)
                return false;

            int index = referenceId.Length - 1;
            while (index >= 0 && char.IsDigit(referenceId[index]))
            {
                index--;
            }

            return int.TryParse(referenceId.Substring(index + 1), out numericPart);
        }

        public void ExtractAndSaveImages (string outputFolder)
        {

            // Delete the output folder if it exists
            if (Directory.Exists(outputFolder))
            {
                Directory.Delete(outputFolder, true);
            }

            // Create the output folder
            Directory.CreateDirectory(outputFolder);

            int imageIndex = 1;

            foreach (SlidePart slidePart in _PresentationDocument.PresentationPart.SlideParts)
            {
                foreach (var picture in slidePart.Slide.Descendants<P.Picture>())
                {
                    Blip blip = picture.Descendants<Blip>().FirstOrDefault();
                    if (blip != null)
                    {
                        string relationshipId = blip.Embed;
                        ImagePart imagePart = (ImagePart)slidePart.GetPartById(relationshipId);

                        // Save the image to the output folder

                        string outputPath = System.IO.Path.Combine(outputFolder, $"Image_{imageIndex++}.{imagePart.Uri.ToString().Split('.').Last()}");
                        try
                        {
                            using (MemoryStream memoryStream = new MemoryStream())
                            {
                                imagePart.GetStream().CopyTo(memoryStream);

                                // Save the content of the image part to the FileStream
                                File.WriteAllBytes(outputPath, memoryStream.ToArray());
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
            }
            _PresentationDocument.Dispose();

        }

        public String GetSlideRelationshipId(SlidePart slidePart)
        {
            return _PresentationPart.GetIdOfPart(slidePart);
        }
        public void SaveAllNotesToTextFile(string filePath)
        {
            // Create or overwrite the file
            using (StreamWriter writer = new StreamWriter(filePath))
            {
                int slideIndex = 1;
                // Iterate through each slide
                foreach (SlidePart slidePart in _PresentationPart.SlideParts)
                {
                    // Check if the slide has notes
                    if (slidePart.NotesSlidePart != null)
                    {
                        // Access the notes slide
                        NotesSlide notesSlide = slidePart.NotesSlidePart.NotesSlide;

                        // Get text from notes slide
                        string noteText = notesSlide.Descendants<D.Text>().Select(t => t.Text).FirstOrDefault();

                        // Write notes text to file
                        writer.WriteLine($"Slide {slideIndex}:");
                        writer.WriteLine(noteText);
                        writer.WriteLine();
                        slideIndex++;
                    }
                }
            }
        }
        public String RemoveSlide (int index)
        {

            // Get the presentation
            P.Presentation presentation = _PresentationPart.Presentation;

            // Get the slide ID list
            SlideIdList slideIdList = presentation.SlideIdList;

            // Get the slide ID to be deleted
            SlideId slideId = slideIdList.Elements<SlideId>().ElementAt(index);

            // Get the relationship ID of the slide
            string slideRelId = slideId.RelationshipId;

            // Remove the slide reference from the slide ID list
            slideIdList.RemoveChild(slideId);

            var slidePart = _PresentationPart.GetPartById(slideRelId);

            // Remove the slide part
            _PresentationPart.DeletePart(slidePart);

            try
            {
                // Remove the relationship reference to the slide part
                _PresentationPart.DeleteReferenceRelationship(slideRelId);
            }
            catch (Exception ex)
            {

            }

            // Save the changes


            return "The slide at index " + index + " has been removed";


        }

        public void AppendSlide (SlideFacade slideFacade)
        {
           
            slideFacade.PresentationSlide.Save(slideFacade.SlidePart);
            _PresentationSlideParts.Add(slideFacade.SlidePart);


        }
        public void Clone(SlideFacade slideFacade)
        {           

            slideFacade.PresentationSlide.Save(slideFacade.SlidePart);
            _PresentationSlideParts.Add(slideFacade.SlidePart);


        }
        public void InsertSlide (int index, SlideFacade slideFacade)
        {
            slideFacade.PresentationSlide.Save(slideFacade.SlidePart);
            _PresentationSlideParts.Add(slideFacade.SlidePart);
            MoveSlideToIndex(slideFacade.SlideIndex, index);

        }

        public void MoveSlideToIndex (int currentIndex, int newIndex)
        {

            // Ensure both indices are valid
            if (currentIndex >= 0 && currentIndex <= _SlideIdList.Count() && newIndex >= 0 && newIndex <= _SlideIdList.Count())
            {
                // Get the SlideId at the current index
                SlideId slideId = _SlideIdList.Elements<SlideId>().ElementAt(currentIndex);

                // Remove the SlideId from its current position
                _SlideIdList.RemoveChild(slideId);

                // Insert the SlideId at the new index
                _SlideIdList.InsertAt(slideId, newIndex);
            }
        }

        public void CopySlide(SlideFacade slideFacade)
        {
            // Open source and target presentations
          
            var targetPresentation = this;

            // Get the source slide part
            SlidePart sourceSlidePart = slideFacade.SlidePart;

            // Get the target presentation part
            PresentationPart targetPresentationPart = getInstance(FilePath).GetPresentationPart();
            SlideIdList targetSlideIdList = targetPresentationPart.Presentation.SlideIdList;
            uint newSlideId = _SlideIdList.Elements<SlideId>().Max(s => s.Id.Value) + 1;


            SlidePart newSlidePart = CopySlidePart(sourceSlidePart, targetPresentationPart);

            SlideId newSlideIdElement = new SlideId()
            {
                Id = newSlideId,
                RelationshipId = _PresentationPart.GetIdOfPart(newSlidePart)
            };
            _SlideIdList.Append(newSlideIdElement);


        }
        public static SlidePart CopySlidePart(SlidePart sourceSlidePart, PresentationPart destinationPresentationPart)
        {
            SlidePart newSlidePart = destinationPresentationPart.AddNewPart<SlidePart>();

            // Clone slide but prevent locking issues
            newSlidePart.Slide = (P.Slide)sourceSlidePart.Slide.CloneNode(true);

            // Handle Slide Layout properly
            if (sourceSlidePart.SlideLayoutPart != null)
            {
                SlideLayoutPart destLayoutPart = destinationPresentationPart.SlideMasterParts
                    .SelectMany(master => master.SlideLayoutParts)
                    .FirstOrDefault(layout => layout.Uri == sourceSlidePart.SlideLayoutPart.Uri);

                if (destLayoutPart != null)
                {
                    newSlidePart.AddPart(destLayoutPart);
                }
            }

            newSlidePart.Slide.Save();
            return newSlidePart;
        }


        public static void Save(string FilePath = null)
        {
            var instance = getInstance(FilePath);
            instance?.Save();
        }

        private void RemoveParts()
        {
            _PresentationDocument = null;
            _PresentationPart = null;
            _PresentationSlideParts = null;
            _CommentAuthorPart = null;
            _PresentationSlideLayoutParts = null;
            _SlideIdList = null;
            _PresentationSlideMasterPart = null;
            
        }
        public void Close(string FilePath = null)
        {
            if (FilePath == null)
            {
                FilePath = _instances.FirstOrDefault(kvp => kvp.Value == _lastInstance).Key;
            }
            if (FilePath != null && _instances.ContainsKey(FilePath))
            {
                _instances[FilePath].Save(); // Ensure saving before closing
                _instances[FilePath].Dispose();
                _instances.Remove(FilePath);
                if (_lastInstance == _instances.GetValueOrDefault(FilePath))
                {
                    _lastInstance = null;
                }
            }
        }

        public void Save()
        {
            if (IsNewPresentation)
            {
                CreatePresentationParts();
            }
            _PresentationDocument?.PresentationPart?.Presentation.Save();
            _PresentationDocument?.Save();
        }
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    _PresentationDocument?.Dispose();
                }
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

    }
}

