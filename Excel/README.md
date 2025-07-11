# Excel Spreadsheets Manipulation In C#

## Basic Workbook Operations
```csharp
// Create a new workbook
var workbook = new Openize.Cells.Workbook();
workbook.Save("excel.xlsx");
Console.WriteLine("Empty excel document created !!!");

// Open an existing workbook
var existingWorkbook = new Openize.Cells.Workbook("existing.xlsx");
```

## Worksheet Management
```csharp
var workbook = new Openize.Cells.Workbook();

// Add a new worksheet
var newSheet = workbook.AddSheet("MyNewSheet");

// Rename a worksheet
workbook.RenameSheet("Sheet1", "RenamedSheet");

// Copy a worksheet
workbook.CopySheet("Sheet1", "CopiedSheet");

// Remove a worksheet
workbook.RemoveSheet("Sheet1");

// Set worksheet visibility
workbook.SetSheetVisibility("Sheet1", SheetVisibility.Hidden);

// Reorder worksheets
workbook.ReorderSheets("Sheet1", 2); // Move to position 2
```

## Cell Operations
```csharp
var workbook = new Openize.Cells.Workbook();
var worksheet = workbook.Worksheets[0];

// Set cell values
worksheet.Cells["A1"].PutValue("Hello World");
worksheet.Cells["B1"].PutValue(42);
worksheet.Cells["C1"].PutValue(3.14);
worksheet.Cells["D1"].PutValue(DateTime.Now);

// Get cell values
string textValue = worksheet.Cells["A1"].GetValue();
int numberValue = worksheet.Cells["B1"].GetValue<int>();

// Set cell styles
worksheet.Cells["A1"].SetFont("Arial", 12, "FF0000"); // Red text
worksheet.Cells["B1"].SetBold(true);
worksheet.Cells["C1"].SetItalic(true);
worksheet.Cells["D1"].SetUnderline(true);
```

## Range Operations
```csharp
var workbook = new Openize.Cells.Workbook();
var worksheet = workbook.Worksheets[0];

// Get a range of cells
var range = worksheet.GetRange("A1:D10");
var range2 = worksheet.GetRange(1, 1, 10, 4); // Row 1, Column 1 to Row 10, Column 4

// Set values for entire range
range.SetValue("Sample Data");

// Copy range to another location
worksheet.CopyRange(range, "F1");

// Merge cells
worksheet.MergeCells("A1", "D1");
```

## Data Import and Export
```csharp
var workbook = new Openize.Cells.Workbook();
var worksheet = workbook.Worksheets[0];

// Import from CSV
var csvOptions = new CsvOptions
{
    Delimiter = ",",
    HasHeader = true,
    Culture = "en-US"
};
int rowsImported = worksheet.ImportFromCsv("data.csv", "A1", csvOptions);

// Export to CSV
int rowsExported = worksheet.ExportToCsv("export.csv", "A1:D10", csvOptions);

// Import from JSON
var jsonOptions = new JsonOptions
{
    IncludeHeaders = true,
    FlattenNestedObjects = true
};
int jsonRowsImported = worksheet.ImportFromJson("data.json", "A1", jsonOptions);

// Export to JSON
int jsonRowsExported = worksheet.ExportToJson("export.json", "A1:D10", jsonOptions);
```

## Row and Column Operations
```csharp
var workbook = new Openize.Cells.Workbook();
var worksheet = workbook.Worksheets[0];

// Row operations
worksheet.SetRowHeight(1, 25.0);
worksheet.HideRow(2);
worksheet.UnhideRow(2);
worksheet.InsertRow(3);
worksheet.InsertRows(4, 3); // Insert 3 rows starting at row 4

// Column operations
worksheet.SetColumnWidth("A", 15.0);
worksheet.AutoFitColumn("B");
worksheet.HideColumn("C");
worksheet.UnhideColumn("C");
worksheet.InsertColumn("D");
worksheet.InsertColumns("E", 3); // Insert 3 columns starting at column E

// Get row and column counts
int rowCount = worksheet.GetRowCount();
int columnCount = worksheet.GetColumnCount();
```

## Freeze Panes
```csharp
var workbook = new Openize.Cells.Workbook();
var worksheet = workbook.Worksheets[0];

// Freeze first row and first column
worksheet.FreezePane(1, 1);

// Get freeze pane information
int frozenRows = worksheet.FreezePanesRow;
int frozenColumns = worksheet.FreezePanesColumn;
```

## Data Validation
```csharp
var workbook = new Openize.Cells.Workbook();
var worksheet = workbook.Worksheets[0];

// Add dropdown list validation
string[] options = { "Option 1", "Option 2", "Option 3" };
worksheet.AddDropdownListValidation("A1", options);

// Apply custom validation rule
var validationRule = new ValidationRule
{
    Type = ValidationType.WholeNumber,
    Operator = ValidationOperator.Between,
    Formula1 = "1",
    Formula2 = "100"
};
worksheet.ApplyValidation("B1", validationRule);

// Get validation rule for a cell
var rule = worksheet.GetValidationRule("A1");
```

## Images
```csharp
var workbook = new Openize.Cells.Workbook();
var worksheet = workbook.Worksheets[0];

// Add image to worksheet
var image = new Image("path/to/image.jpg");
worksheet.AddImage(image, 1, 1, 5, 5); // Start at row 1, col 1, end at row 5, col 5

// Extract images from worksheet
var images = worksheet.ExtractImages();
foreach (var img in images)
{
    Console.WriteLine($"Image: {img.Path}");
}
```

## Comments
```csharp
var workbook = new Openize.Cells.Workbook();
var worksheet = workbook.Worksheets[0];

// Add comment to cell
var comment = new Comment
{
    Text = "This is a comment",
    Author = "John Doe"
};
worksheet.AddComment("A1", comment);
```

## Sheet Protection
```csharp
var workbook = new Openize.Cells.Workbook();
var worksheet = workbook.Worksheets[0];

// Protect worksheet with password
worksheet.ProtectSheet("mypassword");

// Check if sheet is protected
bool isProtected = worksheet.IsProtected();

// Unprotect sheet
worksheet.UnprotectSheet();
```

## Styling and Formatting
```csharp
var workbook = new Openize.Cells.Workbook();
var worksheet = workbook.Worksheets[0];

// Create custom styles
uint customStyleId = workbook.CreateStyle("Arial", 14, "0000FF", 
    HorizontalAlignment.Center, VerticalAlignment.Center);

// Apply style to cell
worksheet.Cells["A1"].StyleId = customStyleId;

// Update default workbook style
workbook.UpdateDefaultStyle("Calibri", 11, "000000");
```

## Built-in Document Properties
```csharp
var workbook = new Openize.Cells.Workbook();

// Set document properties
workbook.BuiltinDocumentProperties.Title = "My Excel Document";
workbook.BuiltinDocumentProperties.Subject = "Sample Data";
workbook.BuiltinDocumentProperties.Creator = "John Doe";
workbook.BuiltinDocumentProperties.Created = DateTime.Now;
``` 