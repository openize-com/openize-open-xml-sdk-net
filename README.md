# Openize.OpenXML SDK for .NET

Openize.OpenXML SDK for .NET is a powerful and easy-to-use wrapper around the OpenXML SDK, enabling developers to work seamlessly with Wordprocessing documents (Docx), Excel spreadsheets (Xlsx), and PowerPoint presentations (Pptx) in their .NET applications.

## Features
- **Word:** Create, read, and manipulate Wordprocessing documents (.docx).
- **Excel:** Create, read, and manipulate Excel spreadsheets (.xlsx).
- **PowerPoint:** Create, read, and manipulate PowerPoint presentations (.pptx).

## Directories
1. **Word:** Contains source code for working with Wordprocessing documents.
2. **Excel:** Contains source code for working with Excel spreadsheets.
3. **PowerPoint:** Contains source code for working with PowerPoint presentations.

## Installation
This project is licensed under the MIT License, so you can freely use it in your projects. To include Openize.OpenXML SDK for .NET in your application, simply clone the repository or add the library to your project via NuGet (coming soon).

## Usage

### Create an Empty Word Document
```csharp
var document = new Openize.Words.Document();
document.Save("word.docx");
Console.WriteLine("Empty word document created !!!");
```

### Create an Empty Excel Spreadsheet
```csharp
var workbook = new Openize.Cells.Workbook();
workbook.Save("excel.xlsx");
Console.WriteLine("Empty excel document created !!!");
```

### Create an Empty PowerPoint Presentation
```csharp
var presentation = Openize.Slides.Presentation.Create("powerpoint.pptx");
presentation.AppendSlide(new Openize.Slides.Slide());
presentation.Save();
Console.WriteLine("Empty Presentation document created !!!");
```

## Contributing
Contributions are welcome! Please fork the repository and submit a pull request. For significant changes, please open an issue first to discuss what you would like to contribute.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
