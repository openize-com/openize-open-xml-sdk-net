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
This project is licensed under the MIT License, so you can freely use it in your projects. To include Openize.OpenXML SDK for .NET in your application, simply clone the repository or add the library to your project via NuGet:
```bash
dotnet add package Openize.OpenXML-SDK
```

## Usage

### Create an Empty Word Document
```csharp
var document = new Openize.Words.Document();
document.Save("word.docx");
Console.WriteLine("Empty word document created !!!");
```
For more details please check [word readme](https://github.com/openize-com/openize-open-xml-sdk-net/blob/main/Word/README.md).

### Create an Empty Excel Spreadsheet
```csharp
var workbook = new Openize.Cells.Workbook();
workbook.Save("excel.xlsx");
Console.WriteLine("Empty excel document created !!!");
```
For more details please check [excel readme](https://github.com/openize-com/openize-open-xml-sdk-net/blob/main/Excel/README.md).

### Create an Empty PowerPoint Presentation
```csharp
var presentation = Openize.Slides.Presentation.Create("powerpoint.pptx");
presentation.AppendSlide(new Openize.Slides.Slide());
presentation.Save();
Console.WriteLine("Empty Presentation document created !!!");
```

## Contributing

Thank you for your interest in contributing to this project! We welcome contributions from the community. To ensure a smooth collaboration, please follow these steps when submitting a pull request:

1. **Fork & Clone** – Fork the repository and clone it to your local machine.
2. **Create a Branch** – Work on a new branch specific to your contribution.
3. **Sign the Contributor License Agreement (CLA)** – Before we can accept your first contribution, you must sign our CLA via [CLA Assistant](https://cla-assistant.io). When you submit your first pull request, you will be prompted to sign the agreement. You can also review the CLA here: [https://cla.openize.com/agreement](https://cla.openize.com/agreement).
4. **Submit a Pull Request (PR)** – Once your changes are ready, submit a PR with a clear description of your contribution.
5. **Review Process** – Our maintainers will review your PR and provide feedback if necessary.

By contributing to this project, you agree to the terms of the CLA and confirm that your contributions comply with the project’s licensing policies.

We appreciate your contributions and look forward to collaborating with you! 

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
