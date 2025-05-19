# Word Documents Manipulation In C#
## Word Paragraphs
```csharp
var doc = new Openize.Words.Document();
var body = new Openize.Words.Body(doc);
var paragraph = new Openize.Words.IElements.Paragraph();
paragraph.AddRun(new Openize.Words.IElements.Run {
    Text = "Word Document Created by Openize.OpenXML-SDK",
        FontFamily="Arial",
        FontSize = 15,
        Bold = true
    });
body.AppendChild(paragraph);
doc.Save("WordOpenize.docx");
```
## Word Tables
```csharp
// Initialize a new word document with the default template
var doc = new Openize.Words.Document();
System.Console.WriteLine("Word Document with default template initialized");

// Initialize the body with the new document
var body = new Openize.Words.Body(doc);
System.Console.WriteLine("Body of the Word Document initialized");

// Get all table styles
var tableStyles = doc.GetElementStyles().TableStyles;
System.Console.WriteLine("Table styles loaded");

// Create Headings Paragraph and append to the body.
foreach (var tableStyle in tableStyles)
{
    var table = new Openize.Words.IElements.Table(5,3);
    table.Style = tableStyle;

    table.Column.Width = 2500;

    var rowNumber = 0;
    var columnNumber = 0;

    var para = new Openize.Words.IElements.Paragraph();
    para.Style = Openize.Words.IElements.Headings.Heading1;
    para.AddRun(new Openize.Words.IElements.Run { 
                Text = $"Table With Style '{tableStyle}' : " 
              });

    body.AppendChild(para);

    foreach (var row in table.Rows)
    {
        rowNumber++;
        foreach(var cell in row.Cells)
        {
            columnNumber++;
            para = new Openize.Words.IElements.Paragraph();
            para.AddRun(new Openize.Words.IElements.Run { 
                                Text = $"Row {rowNumber} Column {columnNumber}"
                                });
            cell.Paragraphs.Add(para);
        }
        columnNumber = 0;
    }
    body.AppendChild(table);
    System.Console.WriteLine($"Table with style {tableStyle} created and appended");
}

// Save the newly created Word Document.
doc.Save($"WordTables.docx");
```
## Word Images
```csharp
using (System.Net.WebClient webClient = new System.Net.WebClient())
{
    webClient.DownloadFile("https://i.imgur.com/V8aRisV.jpg",
                        $"Images/image1.jpg");
    System.Console.WriteLine($"First image downloaded...");
    webClient.DownloadFile("https://i.imgur.com/xrbdI7n.png",
                        $"Images/image2.png");
    System.Console.WriteLine($"Second image downloaded...");
    webClient.DownloadFile("https://i.imgur.com/bqzDqUZ.png",
    $"Images/image3.png");
    System.Console.WriteLine($"Third image downloaded...");
}
// Initialize a new word document with the default template
var doc = new Openize.Words.Document();

// Initialize the body with the new document
var body = new Openize.Words.Body(doc);

// Load images from the specified directory
var imageFiles = System.IO.Directory.GetFiles("Images");
foreach (var imageFile in imageFiles)
{
    // Decode the image with SkiaSharp
    using (var skBMP = SkiaSharp.SKBitmap.Decode(imageFile))
    {
        using (var skIMG = SkiaSharp.SKImage.FromBitmap(skBMP))
        {
            var encoded = skIMG.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
            // Initialize the word document image element
            var img = new Openize.Words.IElements.Image();
            // Load data for the word document image element
            img.ImageData = encoded.ToArray();
            img.Height = 350;
            img.Width = 300;
            // Append image element to the word document
            body.AppendChild(img);
            System.Console.WriteLine($"Image {System.IO.Path.GetFullPath(imageFile)} " +
                  $"added to the word document.");
        }
    }
}
// Save the newly created Word Document.
doc.Save($"WordImages.docx");
```
## Word Shapes
```csharp
// Initialize a new word document with the default template
var doc = new Openize.Words.Document();
System.Console.WriteLine("Word Document with default template initialized");

// Initialize the body with the new document
var body = new Openize.Words.Body(doc);
System.Console.WriteLine("Body of the Word Document initialized");

// Instantiate shape element with hexagone and coordinates/size.
var shape = new Openize.Words.IElements.Shape(100, 100, 400, 400,
    Openize.Words.IElements.ShapeType.Hexagone);
// Add hexagone shape to the word document.
body.AppendChild(shape);
System.Console.WriteLine("Hexagone shape added");

// Save the newly created Word Document.
doc.Save($"WordShape.docx");
```
## Word Groupshapes
```csharp
var doc = new Openize.Words.Document();
var body = new Openize.Words.Body(doc);

var diamond = new Openize.Words.IElements.Shape(0, 0, 200, 200,
                    Openize.Words.IElements.ShapeType.Diamond,
                    Openize.Words.IElements.ShapeFillType.Gradient,
                    new Openize.Words.IElements.ShapeFillColors());

var oval = new Openize.Words.IElements.Shape(300, 0, 200, 200,
                    Openize.Words.IElements.ShapeType.Ellipse,
                    Openize.Words.IElements.ShapeFillType.Pattern,
                    new Openize.Words.IElements.ShapeFillColors());

var groupShape = new Openize.Words.IElements.GroupShape(diamond, oval);
body.AppendChild(groupShape);
doc.Save("WordGroupShape.docx");
```