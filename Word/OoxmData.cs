using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using DF = DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using FF = Openize.Words.IElements;
using Openize.Words;
using System.Linq;
using System.Xml.Linq;
using System.IO;

namespace OpenXML.Words.Data
{
    internal class OoxmlDocData
    {
        private static ConcurrentDictionary<int, WordprocessingDocument> _staticDocDict =
            new ConcurrentDictionary<int, WordprocessingDocument>();
        private static Dictionary<int, WP.Table> _staticTableDict =
            new Dictionary<int, WP.Table>();
        private static int _staticDocCount = 0;
        private OwDocument _ooxmlDoc;
        private readonly object _lockObject = new object();

        private OoxmlDocData(WordprocessingDocument doc)
        {
            lock (_lockObject)
            {
                _ooxmlDoc = OwDocument.CreateInstance(doc);
                _staticDocCount++;
                _staticDocDict.TryAdd(_staticDocCount, doc);
            }
        }

        private OoxmlDocData()
        {
            lock (_lockObject)
            {
                _ooxmlDoc = OwDocument.CreateInstance();
            }
        }

        internal static OoxmlDocData CreateInstance(WordprocessingDocument doc)
        {
            return new OoxmlDocData(doc);
        }

        internal static OoxmlDocData CreateInstance()
        {
            return new OoxmlDocData();
        }

        internal static string ConstructMessage(Exception ex, string operation)
        {
            return $"Error in operation {operation} at OpenXML.Words.Data : {ex.Message} \n Inner Exception: {ex.InnerException?.Message ?? "N/A"}";
        }

        internal static void MapTable(int elementID,WP.Table wpTable)
        {
             _staticTableDict.TryAdd(elementID, wpTable);
        }

        internal void Insert(FF.IElement newElement, int position, Document doc)
        {
            lock (_lockObject)
            {
                _staticDocDict.TryGetValue(doc.GetInstanceInfo(), out WordprocessingDocument staticDoc);

                if (staticDoc?.MainDocumentPart?.Document?.Body == null) throw new OpenizeException("Package or Document or Body is null", new NullReferenceException());

                _ooxmlDoc = OwDocument.CreateInstance(staticDoc);

                var enumerable = staticDoc.MainDocumentPart.Document.Body.Elements().ToList();
                var originalElements = new List<DF.OpenXmlElement>(enumerable);

                var elements = staticDoc.MainDocumentPart.Document.Body.Elements();

                try
                {

                    switch (newElement)
                    {
                        case FF.Paragraph ffPara:
                            //var wpPara = _ooxmlDoc.CreateParagraph(ffPara);
                            var wpPara = OoxmlParagraph.CreateInstance(
                                _ooxmlDoc.IDs, _ooxmlDoc.NumberingPart).CreateParagraph(ffPara);
                            elements.ElementAt(position).InsertBeforeSelf(wpPara);
                            break;

                        case FF.Table ffTable:
                            var wpTable = OoxmlTable.CreateInstance(
                                _ooxmlDoc.IDs, _ooxmlDoc.NumberingPart).CreateTable(ffTable);
                            elements.ElementAt(position).InsertBeforeSelf(wpTable);
                            break;

                        case FF.Image ffImage:
                            //var wpImage = _ooxmlDoc.CreateImage(ffImage, staticDoc.MainDocumentPart);
                            var wpImage = OoxmlImage.CreateInstance(
                                _ooxmlDoc.MainPart).CreateImage(ffImage);
                            elements.ElementAt(position).InsertBeforeSelf(wpImage);
                            break;

                        case FF.Shape ffShape:
                            //var wpShape = _ooxmlDoc.CreateShape(ffShape);
                            var wpShape = OoxmlShape.CreateInstance().CreateShape(ffShape);
                            elements.ElementAt(position).InsertBeforeSelf(wpShape);
                            break;

                        case FF.GroupShape ffGroupShape:
                            //var wpGroupShape = _ooxmlDoc.CreateGroupShape(ffGroupShape);
                            var wpGroupShape = OoxmlGroupShape.CreateInstance().
                                CreateGroupShape(ffGroupShape);
                            elements.ElementAt(position).InsertBeforeSelf(wpGroupShape);
                            break;
                    }

                }
                catch (Exception ex)
                {
                    staticDoc.MainDocumentPart.Document.Body.RemoveAllChildren();
                    staticDoc.MainDocumentPart.Document.Body.Append(originalElements);
                    var errorMessage = ConstructMessage(ex, "Insert OOXML Element(s)");
                    throw new OpenizeException(errorMessage, ex);
                }
            }
        }

        internal void Update(FF.IElement newElement, int position, Document doc)
        {
            lock (_lockObject)
            {
                _staticDocDict.TryGetValue(doc.GetInstanceInfo(), out WordprocessingDocument staticDoc);

                if (staticDoc?.MainDocumentPart?.Document?.Body == null) throw new OpenizeException("Package or Document or Body is null", new NullReferenceException());

                _ooxmlDoc = OwDocument.CreateInstance(staticDoc);

                var enumerable = staticDoc.MainDocumentPart.Document.Body.Elements().ToList();
                var originalElements = new List<DF.OpenXmlElement>(enumerable);

                try
                {
                    if (position >= 0)
                    {
                        var elements = staticDoc.MainDocumentPart.Document.Body.Elements();
                        elements.ElementAt(position).Remove();
                        var enumerable1 = elements.ToList();
                        var existingElement = enumerable1.ElementAt(position);

                        switch (newElement)
                        {
                            case FF.Paragraph ffPara:
                                //var wpPara = _ooxmlDoc.CreateParagraph(ffPara);
                                var wpPara = OoxmlParagraph.CreateInstance(
                                   _ooxmlDoc.IDs, _ooxmlDoc.NumberingPart).CreateParagraph(ffPara);
                                enumerable1.ElementAt(position).InsertBeforeSelf(wpPara);
                                break;

                            case FF.Table ffTable:
                                _staticTableDict.TryGetValue(ffTable.ElementId, out WP.Table wpOldTable);
                                //var wpTable = OoxmlTable.CreateInstance(
                                //   _ooxmlDoc.IDs, _ooxmlDoc.NumberingPart).CreateTable(ffTable);
                                var wpTable = OoxmlTable.CreateInstance(
                                    _ooxmlDoc.IDs, _ooxmlDoc.NumberingPart).
                                    UpdateTable(ffTable, wpOldTable);
                                enumerable1.ElementAt(position).InsertBeforeSelf(wpTable);
                                break;
                            case FF.Image ffImage:
                                //var wpImage = _ooxmlDoc.CreateImage(ffImage, staticDoc.MainDocumentPart);
                                var wpImage = OoxmlImage.CreateInstance(
                                   _ooxmlDoc.MainPart).CreateImage(ffImage);
                                enumerable1.ElementAt(position).InsertBeforeSelf(wpImage);
                                break;

                            case FF.Shape ffShape:
                                //var wpShape = _ooxmlDoc.CreateShape(ffShape);
                                var wpShape = OoxmlShape.CreateInstance().CreateShape(ffShape);
                                enumerable1.ElementAt(position).InsertBeforeSelf(wpShape);
                                break;

                            case FF.GroupShape ffGroupShape:
                                //var wpGroupShape = _ooxmlDoc.CreateGroupShape(ffGroupShape);
                                var wpGroupShape = OoxmlGroupShape.CreateInstance().
                                    CreateGroupShape(ffGroupShape);
                                elements.ElementAt(position).InsertBeforeSelf(wpGroupShape);
                                break;
                        }

                    }
                }
                catch (Exception ex)
                {
                    // Rollback changes by restoring the original elements
                    staticDoc.MainDocumentPart.Document.Body.RemoveAllChildren();
                    staticDoc.MainDocumentPart.Document.Body.Append(originalElements);
                    var errorMessage = ConstructMessage(ex, "Update OOXML Element(s)");
                    throw new OpenizeException(errorMessage, ex);
                }
            }
        }

        internal void Remove(int position, Document doc)
        {
            lock (_lockObject)
            {
                _staticDocDict.TryGetValue(doc.GetInstanceInfo(), out WordprocessingDocument staticDoc);

                if (staticDoc?.MainDocumentPart?.Document?.Body == null) throw new OpenizeException("Package or Document or Body is null", new NullReferenceException());

                var enumerable = staticDoc.MainDocumentPart.Document.Body.Elements().ToList();
                var originalElements = new List<DF.OpenXmlElement>(enumerable);

                try
                {
                    var elements = staticDoc.MainDocumentPart.Document.Body.Elements();
                    elements.ElementAt(position).Remove();
                }
                catch (Exception ex)
                {
                    // Rollback changes by restoring the original elements
                    staticDoc.MainDocumentPart.Document.Body.RemoveAllChildren();
                    staticDoc.MainDocumentPart.Document.Body.Append(originalElements);
                    var errorMessage = ConstructMessage(ex, "Remove OOXML Element(s)");
                    throw new OpenizeException(errorMessage, ex);
                }
            }
        }

        internal void Append(FF.IElement newElement, Document doc)
        {
            lock (_lockObject)
            {
                _staticDocDict.TryGetValue(doc.GetInstanceInfo(), out WordprocessingDocument staticDoc);

                if (staticDoc?.MainDocumentPart?.Document?.Body == null) throw new OpenizeException("Package or Document or Body is null", new NullReferenceException());

                _ooxmlDoc = OwDocument.CreateInstance(staticDoc);

                var enumerable = staticDoc.MainDocumentPart.Document.Body.Elements().ToList();
                var originalElements = new List<DF.OpenXmlElement>(enumerable);

                var sectionPropertiesList = staticDoc.MainDocumentPart.Document.Body.Elements<WP.SectionProperties>().ToList();
                WP.SectionProperties lastSectionProperties = null;
                if (sectionPropertiesList.Any()) lastSectionProperties = sectionPropertiesList.Last();

                try
                {

                    switch (newElement)
                    {
                        case FF.Paragraph ffPara:
                            //var wpPara = _ooxmlDoc.CreateParagraph(ffPara);
                            var wpPara = OoxmlParagraph.CreateInstance(
                                _ooxmlDoc.IDs, _ooxmlDoc.NumberingPart).CreateParagraph(ffPara);
                            if (lastSectionProperties != null) staticDoc.MainDocumentPart.Document.Body.InsertBefore(wpPara, lastSectionProperties);
                            else staticDoc.MainDocumentPart.Document.Body.Append(wpPara);
                            break;
                        case FF.Table ffTable:
                            var wpTable = OoxmlTable.CreateInstance(
                                _ooxmlDoc.IDs, _ooxmlDoc.NumberingPart).CreateTable(ffTable);
                            if (lastSectionProperties != null) staticDoc.MainDocumentPart.Document.Body.InsertBefore(wpTable, lastSectionProperties);
                            else staticDoc.MainDocumentPart.Document.Body.Append(wpTable);
                            break;
                        case FF.Image ffImage:
                            //var wpImage = _ooxmlDoc.CreateImage(ffImage, staticDoc.MainDocumentPart);
                            var wpImage = OoxmlImage.CreateInstance(
                                   _ooxmlDoc.MainPart).CreateImage(ffImage);
                            if (lastSectionProperties != null) staticDoc.MainDocumentPart.Document.Body.InsertBefore(wpImage, lastSectionProperties);
                            else staticDoc.MainDocumentPart.Document.Body.Append(wpImage);
                            break;
                        case FF.Shape ffShape:
                            //var wpShape = _ooxmlDoc.CreateShape(ffShape);
                            var wpShape = OoxmlShape.CreateInstance().CreateShape(ffShape);
                            if (lastSectionProperties != null) staticDoc.MainDocumentPart.Document.Body.InsertBefore(wpShape, lastSectionProperties);
                            else staticDoc.MainDocumentPart.Document.Body.Append(wpShape);
                            break;

                        case FF.GroupShape ffGroupShape:
                            //var wpGroupShape = _ooxmlDoc.CreateGroupShape(ffGroupShape);
                            var wpGroupShape = OoxmlGroupShape.CreateInstance().
                                CreateGroupShape(ffGroupShape);
                            if (lastSectionProperties != null) staticDoc.MainDocumentPart.Document.Body.InsertBefore(wpGroupShape, lastSectionProperties);
                            else staticDoc.MainDocumentPart.Document.Body.Append(wpGroupShape);
                            break;
                    }

                }
                catch (Exception ex)
                {
                    // Rollback changes by restoring the original elements
                    staticDoc.MainDocumentPart.Document.Body.RemoveAllChildren();
                    staticDoc.MainDocumentPart.Document.Body.Append(originalElements);
                    var errorMessage = ConstructMessage(ex, "Append OOXML Element(s)");
                    throw new OpenizeException(errorMessage, ex);
                }
            }
        }

        /**
        internal void UpdateProperties(Document doc)
        {
            _staticDocDict.TryGetValue(doc.GetInstanceInfo(), out WordprocessingDocument staticDoc);
            DocumentProperties documentProperties = doc.GetDocumentProperties();
            staticDoc.Save(); 
            var corePart = staticDoc.CoreFilePropertiesPart;
            XDocument coreXml;

            using (MemoryStream memoryStream = new MemoryStream())
            {
                // Read the core properties XML into memory
                using (Stream readStream = corePart.GetStream(FileMode.Open, FileAccess.Read))
                {
                    readStream.CopyTo(memoryStream);
                }

                memoryStream.Position = 0;
                coreXml = XDocument.Load(memoryStream);
            }

            // Define the XML namespace for Dublin Core metadata
            XNamespace dc = "http://purl.org/dc/elements/1.1/";
            XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
            XNamespace dcterms = "http://purl.org/dc/terms/";

            // Find or create the <dc:title> element
            var titleElement = coreXml.Descendants(dc + "title").FirstOrDefault();
            if (titleElement != null)
            {
                titleElement.Value = documentProperties.Title; // Update existing title
            }
            else
            {
                coreXml.Root.Add(new XElement(dc + "title", documentProperties.Title)); // Add title if missing
            }

            // Find or create the <dc:subject> element
            var subjectElement = coreXml.Descendants(dc + "subject").FirstOrDefault();
            if (subjectElement != null)
            {
                subjectElement.Value = documentProperties.Subject; // Update existing subject
            }
            else
            {
                coreXml.Root.Add(new XElement(dc + "subject", documentProperties.Subject)); // Add subject if missing
            }

            // Find or create the <dc:creator> element
            var creatorElement = coreXml.Descendants(dc + "creator").FirstOrDefault();
            if (creatorElement != null)
            {
                creatorElement.Value = documentProperties.Creator; // Update existing creator
            }
            else
            {
                coreXml.Root.Add(new XElement(dc + "creator", documentProperties.Creator)); // Add creator if missing
            }

            // Find or create the <cp:keywords> element
            var keywordsElement = coreXml.Descendants(cp + "keywords").FirstOrDefault();
            if (keywordsElement != null)
            {
                keywordsElement.Value = documentProperties.Keywords; // Update existing keywords
            }
            else
            {
                coreXml.Root.Add(new XElement(cp + "keywords", documentProperties.Keywords)); // Add keywords if missing
            }

            // Save the updated XML back to the CoreFilePropertiesPart
            using (Stream writeStream = corePart.GetStream(FileMode.Create, FileAccess.Write))
            {
                coreXml.Save(writeStream);
            }

            staticDoc.Save(); // Save the document after updating properties
        }
        **/

        internal void UpdateProperties(Document doc)
        {
            _staticDocDict.TryGetValue(doc.GetInstanceInfo(), out WordprocessingDocument staticDoc);
            DocumentProperties documentProperties = doc.GetDocumentProperties();

            staticDoc.Save();
            var corePart = staticDoc.CoreFilePropertiesPart;

            XDocument coreXml;
            using (var memoryStream = new MemoryStream())
            {
                using (var readStream = corePart.GetStream(FileMode.Open, FileAccess.Read))
                {
                    readStream.CopyTo(memoryStream);
                }

                memoryStream.Position = 0;
                coreXml = XDocument.Load(memoryStream);
            }

            XNamespace dc = "http://purl.org/dc/elements/1.1/";
            XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
            XNamespace dcterms = "http://purl.org/dc/terms/";

            // Helper to update or insert a value
            void UpdateElement(XNamespace ns, string name, string value)
            {
                if (string.IsNullOrWhiteSpace(value)) return;

                var element = coreXml.Descendants(ns + name).FirstOrDefault();
                if (element != null)
                    element.Value = value;
                else
                    coreXml.Root.Add(new XElement(ns + name, value));
            }

            UpdateElement(dc, "title", documentProperties.Title);
            UpdateElement(dc, "subject", documentProperties.Subject);
            UpdateElement(dc, "description", documentProperties.Description);
            UpdateElement(dc, "creator", documentProperties.Creator);
            UpdateElement(cp, "keywords", documentProperties.Keywords);
            UpdateElement(cp, "lastModifiedBy", documentProperties.LastModifiedBy);
            UpdateElement(cp, "revision", documentProperties.Revision);
            UpdateElement(dcterms, "created", documentProperties.Created);
            UpdateElement(dcterms, "modified", documentProperties.Modified);

            using (var writeStream = corePart.GetStream(FileMode.Create, FileAccess.Write))
            {
                coreXml.Save(writeStream);
            }

            staticDoc.Save();
        }


        internal void Save(System.IO.Stream stream, Document doc)
        {
            lock (_lockObject)
            {
                try
                {

                    _staticDocDict.TryGetValue(doc.GetInstanceInfo(), out WordprocessingDocument staticDoc);

                    staticDoc.Clone(stream);
                }
                catch (Exception ex)
                {
                    var errorMessage = ConstructMessage(ex, "Save OOXML OWDocument");
                    throw new OpenizeException(errorMessage, ex);
                }
            }
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_ooxmlDoc == null) return;
                _ooxmlDoc.Dispose();
                _ooxmlDoc = null;
            }
            // Dispose of unmanaged resources
        }
    }
}
