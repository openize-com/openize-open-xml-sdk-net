using System;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Openize.Cells
{
    /// <summary>
    /// Extension methods for data import and export operations on worksheets.
    /// </summary>
    public static class WorksheetDataExtensions
    {
        /// <summary>
        /// Imports data from a CSV file into the worksheet starting at the specified cell.
        /// </summary>
        /// <param name="worksheet">The worksheet to import data into.</param>
        /// <param name="filePath">Path to the CSV file.</param>
        /// <param name="startCellReference">Starting cell reference (e.g., "A1").</param>
        /// <param name="options">CSV import options. If null, default options will be used.</param>
        /// <returns>The number of rows imported.</returns>
        /// <exception cref="FileNotFoundException">Thrown when the CSV file is not found.</exception>
        /// <exception cref="ArgumentException">Thrown when parameters are invalid.</exception>
        public static int ImportFromCsv(this Worksheet worksheet, string filePath, string startCellReference = "A1", CsvOptions options = null)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("CSV file not found.", filePath);

            if (string.IsNullOrEmpty(startCellReference))
                throw new ArgumentException("Start cell reference cannot be null or empty.", nameof(startCellReference));

            // Use default options if none provided
            options = options ?? new CsvOptions();

            var (startRow, startColumn) = ParseCellReference(startCellReference);
            var culture = new CultureInfo(options.Culture);

            string[] lines;

            // Read file with specified encoding
            using (var reader = new StreamReader(filePath, options.Encoding))
            {
                var content = reader.ReadToEnd();
                lines = content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            }

            if (options.SkipEmptyLines)
            {
                lines = lines.Where(line => !string.IsNullOrWhiteSpace(line)).ToArray();
            }

            int rowsImported = 0;
            int maxRowsToImport = options.MaxRows > 0 ? Math.Min(options.MaxRows, lines.Length) : lines.Length;

            for (int i = 0; i < maxRowsToImport; i++)
            {
                var line = lines[i];
                var values = ParseCsvLine(line, options.Delimiter, options.TextQualifier);

                for (int j = 0; j < values.Length; j++)
                {
                    var cellRef = $"{IndexToColumnLetter((int)(startColumn + j - 1))}{startRow + i}";
                    var cell = worksheet.GetCell(cellRef);

                    var value = values[j];

                    if (options.TrimWhitespace)
                    {
                        value = value?.Trim();
                    }

                    if (string.IsNullOrEmpty(value))
                    {
                        cell.PutValue("");
                        continue;
                    }

                    // Auto-detect data types if enabled
                    if (options.AutoDetectDataTypes)
                    {
                        if (DateTime.TryParseExact(value, options.DateFormat, culture, DateTimeStyles.None, out DateTime dateValue))
                        {
                            cell.PutValue(dateValue);
                        }
                        else if (double.TryParse(value, NumberStyles.Any, culture, out double doubleValue))
                        {
                            cell.PutValue(doubleValue);
                        }
                        else
                        {
                            cell.PutValue(value);
                        }
                    }
                    else
                    {
                        cell.PutValue(value);
                    }
                }

                rowsImported++;
            }

            return rowsImported;
        }

        /// <summary>
        /// Exports worksheet data to a CSV file.
        /// </summary>
        /// <param name="worksheet">The worksheet to export data from.</param>
        /// <param name="filePath">Path where the CSV file will be saved.</param>
        /// <param name="range">Range to export (e.g., "A1:E10"). If null, exports all used range.</param>
        /// <param name="options">CSV export options. If null, default options will be used.</param>
        /// <returns>The number of rows exported.</returns>
        public static int ExportToCsv(this Worksheet worksheet, string filePath, string range = null, CsvOptions options = null)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));

            // Use default options if none provided
            options = options ?? new CsvOptions();

            // Determine export range
            Range exportRange = DetermineExportRange(worksheet, range);

            var culture = new CultureInfo(options.Culture);
            var lines = new List<string>();

            for (uint row = exportRange.StartRowIndex; row <= exportRange.EndRowIndex; row++)
            {
                var values = new List<string>();

                for (uint col = exportRange.StartColumnIndex; col <= exportRange.EndColumnIndex; col++)
                {
                    var cellRef = $"{IndexToColumnLetter((int)col)}{row}";
                    var cell = worksheet.GetCell(cellRef);
                    var value = cell.GetValue() ?? "";

                    // Handle text qualifier if value contains delimiter or qualifier
                    if (value.Contains(options.Delimiter) || value.Contains(options.TextQualifier))
                    {
                        value = $"{options.TextQualifier}{value.Replace(options.TextQualifier, options.TextQualifier + options.TextQualifier)}{options.TextQualifier}";
                    }

                    values.Add(value);
                }

                lines.Add(string.Join(options.Delimiter, values));
            }

            // Write to file with specified encoding
            using (var writer = new StreamWriter(filePath, false, options.Encoding))
            {
                foreach (var line in lines)
                {
                    writer.WriteLine(line);
                }
            }

            return lines.Count;
        }

        /// <summary>
        /// Imports data from a JSON file into the worksheet starting at the specified cell.
        /// Supports both JSON arrays and single JSON objects.
        /// </summary>
        /// <param name="worksheet">The worksheet to import data into.</param>
        /// <param name="filePath">Path to the JSON file.</param>
        /// <param name="startCellReference">Starting cell reference (e.g., "A1").</param>
        /// <param name="options">JSON import options. If null, default options will be used.</param>
        /// <returns>The number of rows imported.</returns>
        /// <exception cref="FileNotFoundException">Thrown when the JSON file is not found.</exception>
        /// <exception cref="ArgumentException">Thrown when parameters are invalid.</exception>
        /// <exception cref="JsonException">Thrown when JSON parsing fails.</exception>
        public static int ImportFromJson(this Worksheet worksheet, string filePath, string startCellReference = "A1", JsonOptions options = null)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("JSON file not found.", filePath);

            var jsonContent = File.ReadAllText(filePath);
            return worksheet.ImportFromJsonString(jsonContent, startCellReference, options);
        }

        /// <summary>
        /// Imports data from a JSON string into the worksheet starting at the specified cell.
        /// </summary>
        /// <param name="worksheet">The worksheet to import data into.</param>
        /// <param name="jsonString">The JSON string containing the data.</param>
        /// <param name="startCellReference">Starting cell reference (e.g., "A1").</param>
        /// <param name="options">JSON import options. If null, default options will be used.</param>
        /// <returns>The number of rows imported.</returns>
        public static int ImportFromJsonString(this Worksheet worksheet, string jsonString, string startCellReference = "A1", JsonOptions options = null)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            if (string.IsNullOrEmpty(jsonString))
                throw new ArgumentException("JSON string cannot be null or empty.", nameof(jsonString));

            if (string.IsNullOrEmpty(startCellReference))
                throw new ArgumentException("Start cell reference cannot be null or empty.", nameof(startCellReference));

            // Use default options if none provided
            options = options ?? new JsonOptions();

            var (startRow, startColumn) = ParseCellReference(startCellReference);
            var culture = new CultureInfo(options.Culture);

            try
            {
                var jsonToken = JToken.Parse(jsonString);
                var records = new List<Dictionary<string, object>>();

                if (jsonToken is JArray jsonArray)
                {
                    // Handle JSON array
                    foreach (var item in jsonArray)
                    {
                        if (item is JObject jObj)
                        {
                            var record = ProcessJsonObject(jObj, options);
                            records.Add(record);
                        }
                    }
                }
                else if (jsonToken is JObject jsonObject)
                {
                    // Handle single JSON object
                    var record = ProcessJsonObject(jsonObject, options);
                    records.Add(record);
                }
                else
                {
                    throw new JsonException("JSON must be either an object or an array of objects.");
                }

                if (!records.Any())
                {
                    return 0;
                }

                // Apply max records limit
                if (options.MaxRecords > 0 && records.Count > options.MaxRecords)
                {
                    records = records.Take(options.MaxRecords).ToList();
                }

                // Get all unique column names
                var allColumns = records
                    .SelectMany(r => r.Keys)
                    .Distinct()
                    .OrderBy(k => k)
                    .ToList();

                int currentRow = (int)startRow;
                int rowsImported = 0;

                // Add headers if enabled
                if (options.IncludeHeaders)
                {
                    for (int col = 0; col < allColumns.Count; col++)
                    {
                        var cellRef = $"{IndexToColumnLetter((int)(startColumn + col - 1))}{currentRow}";
                        var cell = worksheet.GetCell(cellRef);
                        cell.PutValue(allColumns[col]);
                    }
                    currentRow++;
                    rowsImported++;
                }

                // Add data rows
                foreach (var record in records)
                {
                    for (int col = 0; col < allColumns.Count; col++)
                    {
                        var columnName = allColumns[col];
                        var cellRef = $"{IndexToColumnLetter((int)(startColumn + col - 1))}{currentRow}";
                        var cell = worksheet.GetCell(cellRef);

                        if (record.TryGetValue(columnName, out var value) && value != null)
                        {
                            SetCellValueFromJson(cell, value, options, culture);
                        }
                        else if (!options.IgnoreNullValues)
                        {
                            cell.PutValue("");
                        }
                    }
                    currentRow++;
                    rowsImported++;
                }

                return rowsImported;
            }
            catch (JsonException ex)
            {
                throw new JsonException($"Failed to parse JSON: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Exports worksheet data to a JSON file.
        /// </summary>
        /// <param name="worksheet">The worksheet to export data from.</param>
        /// <param name="filePath">Path where the JSON file will be saved.</param>
        /// <param name="range">Range to export (e.g., "A1:E10"). If null, exports all used range.</param>
        /// <param name="options">JSON export options. If null, default options will be used.</param>
        /// <returns>The number of records exported.</returns>
        public static int ExportToJson(this Worksheet worksheet, string filePath, string range = null, JsonOptions options = null)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));

            // Use default options if none provided
            options = options ?? new JsonOptions();

            // Determine export range
            Range exportRange = DetermineExportRange(worksheet, range);

            var records = new List<Dictionary<string, object>>();
            var headers = new List<string>();

            // Get headers from first row if enabled
            if (options.IncludeHeaders)
            {
                for (uint col = exportRange.StartColumnIndex; col <= exportRange.EndColumnIndex; col++)
                {
                    var cellRef = $"{IndexToColumnLetter((int)col)}{exportRange.StartRowIndex}";
                    var headerValue = worksheet.GetCell(cellRef).GetValue();
                    headers.Add(string.IsNullOrEmpty(headerValue) ? $"Column{col}" : headerValue);
                }
            }
            else
            {
                // Generate default column names
                for (uint col = exportRange.StartColumnIndex; col <= exportRange.EndColumnIndex; col++)
                {
                    headers.Add($"Column{col}");
                }
            }

            // Start from second row if headers are included, otherwise from first row
            uint startDataRow = options.IncludeHeaders ? exportRange.StartRowIndex + 1 : exportRange.StartRowIndex;

            // Extract data rows
            for (uint row = startDataRow; row <= exportRange.EndRowIndex; row++)
            {
                var record = new Dictionary<string, object>();
                bool hasData = false;

                for (uint col = exportRange.StartColumnIndex; col <= exportRange.EndColumnIndex; col++)
                {
                    var cellRef = $"{IndexToColumnLetter((int)col)}{row}";
                    var cellValue = worksheet.GetCell(cellRef).GetValue();
                    var headerName = headers[(int)(col - exportRange.StartColumnIndex)];

                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        hasData = true;
                        record[headerName] = ConvertCellValueForJson(cellValue, options);
                    }
                    else if (!options.IgnoreNullValues)
                    {
                        record[headerName] = null;
                    }
                }

                // Only add record if it has data or we're not ignoring empty rows
                if (hasData || !options.IgnoreNullValues)
                {
                    records.Add(record);
                }

                // Apply max records limit
                if (options.MaxRecords > 0 && records.Count >= options.MaxRecords)
                {
                    break;
                }
            }

            // Convert to JSON and save
            var jsonSettings = new JsonSerializerSettings
            {
                Formatting = Formatting.Indented,
                NullValueHandling = options.IgnoreNullValues ? NullValueHandling.Ignore : NullValueHandling.Include,
                DateFormatString = options.DateTimeFormat
            };

            var json = JsonConvert.SerializeObject(records, jsonSettings);
            File.WriteAllText(filePath, json, Encoding.UTF8);

            return records.Count;
        }

        /// <summary>
        /// Exports worksheet data to a JSON string.
        /// </summary>
        /// <param name="worksheet">The worksheet to export data from.</param>
        /// <param name="range">Range to export (e.g., "A1:E10"). If null, exports all used range.</param>
        /// <param name="options">JSON export options. If null, default options will be used.</param>
        /// <returns>The JSON string containing the exported data.</returns>
        public static string ExportToJsonString(this Worksheet worksheet, string range = null, JsonOptions options = null)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            // Use default options if none provided
            options = options ?? new JsonOptions();

            // Determine export range
            Range exportRange = DetermineExportRange(worksheet, range);

            var records = new List<Dictionary<string, object>>();
            var headers = new List<string>();

            // Get headers from first row if enabled
            if (options.IncludeHeaders)
            {
                for (uint col = exportRange.StartColumnIndex; col <= exportRange.EndColumnIndex; col++)
                {
                    var cellRef = $"{IndexToColumnLetter((int)col)}{exportRange.StartRowIndex}";
                    var headerValue = worksheet.GetCell(cellRef).GetValue();
                    headers.Add(string.IsNullOrEmpty(headerValue) ? $"Column{col}" : headerValue);
                }
            }
            else
            {
                // Generate default column names
                for (uint col = exportRange.StartColumnIndex; col <= exportRange.EndColumnIndex; col++)
                {
                    headers.Add($"Column{col}");
                }
            }

            // Start from second row if headers are included, otherwise from first row
            uint startDataRow = options.IncludeHeaders ? exportRange.StartRowIndex + 1 : exportRange.StartRowIndex;

            // Extract data rows
            for (uint row = startDataRow; row <= exportRange.EndRowIndex; row++)
            {
                var record = new Dictionary<string, object>();
                bool hasData = false;

                for (uint col = exportRange.StartColumnIndex; col <= exportRange.EndColumnIndex; col++)
                {
                    var cellRef = $"{IndexToColumnLetter((int)col)}{row}";
                    var cellValue = worksheet.GetCell(cellRef).GetValue();
                    var headerName = headers[(int)(col - exportRange.StartColumnIndex)];

                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        hasData = true;
                        record[headerName] = ConvertCellValueForJson(cellValue, options);
                    }
                    else if (!options.IgnoreNullValues)
                    {
                        record[headerName] = null;
                    }
                }

                // Only add record if it has data or we're not ignoring empty rows
                if (hasData || !options.IgnoreNullValues)
                {
                    records.Add(record);
                }

                // Apply max records limit
                if (options.MaxRecords > 0 && records.Count >= options.MaxRecords)
                {
                    break;
                }
            }

            // Convert to JSON and return
            var jsonSettings = new JsonSerializerSettings
            {
                Formatting = Formatting.Indented,
                NullValueHandling = options.IgnoreNullValues ? NullValueHandling.Ignore : NullValueHandling.Include,
                DateFormatString = options.DateTimeFormat
            };

            return JsonConvert.SerializeObject(records, jsonSettings);
        }

        private static object ConvertCellValueForJson(string cellValue, JsonOptions options)
        {
            if (string.IsNullOrEmpty(cellValue))
                return null;

            if (!options.AutoDetectDataTypes)
                return cellValue;

            var culture = new CultureInfo(options.Culture);

            // Try to parse as boolean
            if (bool.TryParse(cellValue, out bool boolValue))
                return boolValue;

            // Try to parse as DateTime
            if (DateTime.TryParseExact(cellValue, options.DateTimeFormat, culture, DateTimeStyles.None, out DateTime dateTimeValue))
                return dateTimeValue.ToString(options.DateTimeFormat);

            if (DateTime.TryParseExact(cellValue, options.DateFormat, culture, DateTimeStyles.None, out DateTime dateValue))
                return dateValue.ToString(options.DateFormat);

            // Try to parse as number
            if (double.TryParse(cellValue, NumberStyles.Any, culture, out double doubleValue))
            {
                // Check if it's actually an integer
                if (doubleValue == Math.Floor(doubleValue) && doubleValue >= int.MinValue && doubleValue <= int.MaxValue)
                    return (int)doubleValue;
                return doubleValue;
            }

            // Return as string
            return cellValue;
        }

        /// <summary>
        /// Exports worksheet data to a CSV string.
        /// </summary>
        /// <param name="worksheet">The worksheet to export data from.</param>
        /// <param name="range">Range to export (e.g., "A1:E10"). If null, exports all used range.</param>
        /// <param name="options">CSV export options. If null, default options will be used.</param>
        /// <returns>The CSV string containing the exported data.</returns>
        public static string ExportToCsvString(this Worksheet worksheet, string range = null, CsvOptions options = null)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            // Use default options if none provided
            options = options ?? new CsvOptions();

            // Determine export range
            Range exportRange = DetermineExportRange(worksheet, range);

            var culture = new CultureInfo(options.Culture);
            var lines = new List<string>();
            int recordsProcessed = 0;

            // Start from second row if headers should be skipped, otherwise from first row
            uint startDataRow = (!options.HasHeaders && exportRange.StartRowIndex == 1) ? exportRange.StartRowIndex + 1 : exportRange.StartRowIndex;

            for (uint row = startDataRow; row <= exportRange.EndRowIndex; row++)
            {
                var values = new List<string>();
                bool hasData = false;

                for (uint col = exportRange.StartColumnIndex; col <= exportRange.EndColumnIndex; col++)
                {
                    var cellRef = $"{IndexToColumnLetter((int)col)}{row}";
                    var cell = worksheet.GetCell(cellRef);
                    var value = cell.GetValue() ?? "";

                    if (!string.IsNullOrEmpty(value))
                    {
                        hasData = true;
                    }

                    // Apply trimming if enabled
                    if (options.TrimWhitespace)
                    {
                        value = value.Trim();
                    }

                    // Handle text qualifier if value contains delimiter or qualifier
                    if (value.Contains(options.Delimiter) || value.Contains(options.TextQualifier) || value.Contains("\n") || value.Contains("\r"))
                    {
                        value = $"{options.TextQualifier}{value.Replace(options.TextQualifier, options.TextQualifier + options.TextQualifier)}{options.TextQualifier}";
                    }

                    values.Add(value);
                }

                // Only add line if it has data or we're not skipping empty lines
                if (hasData || !options.SkipEmptyLines)
                {
                    lines.Add(string.Join(options.Delimiter, values));
                    recordsProcessed++;
                }

                // Apply max records limit
                if (options.MaxRows > 0 && recordsProcessed >= options.MaxRows)
                {
                    break;
                }
            }

            return string.Join(Environment.NewLine, lines);
        }

        private static object ConvertCellValueForCsv(string cellValue, CsvOptions options)
        {
            if (string.IsNullOrEmpty(cellValue))
                return string.Empty;

            var culture = new CultureInfo(options.Culture);

            if (!options.AutoDetectDataTypes)
                return cellValue;

            // Try to parse as DateTime and format it
            if (DateTime.TryParseExact(cellValue, options.DateFormat, culture, DateTimeStyles.None, out DateTime dateValue))
                return dateValue.ToString(options.DateFormat, culture);

            // Try to parse as number and format it
            if (double.TryParse(cellValue, NumberStyles.Any, culture, out double doubleValue))
                return doubleValue.ToString(culture);

            // Return as string
            return cellValue;
        }

        #region Private Helper Methods

        private static (uint row, uint column) ParseCellReference(string cellReference)
        {
            var match = Regex.Match(cellReference, @"([A-Z]+)(\d+)");
            if (!match.Success)
                throw new FormatException("Invalid cell reference format.");

            uint row = uint.Parse(match.Groups[2].Value);
            uint column = (uint)ColumnLetterToIndex(match.Groups[1].Value);

            return (row, column);
        }

        private static int ColumnLetterToIndex(string column)
        {
            int index = 0;
            foreach (var ch in column)
            {
                index = (index * 26) + (ch - 'A' + 1);
            }
            return index;
        }

        private static string IndexToColumnLetter(int index)
        {
            string columnLetter = string.Empty;
            while (index >= 0)
            {
                columnLetter = (char)('A' + (index % 26)) + columnLetter;
                index = (index / 26) - 1;
            }
            return columnLetter;
        }

        private static string[] ParseCsvLine(string line, string delimiter, string textQualifier)
        {
            var values = new List<string>();
            var currentValue = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c.ToString() == textQualifier)
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1].ToString() == textQualifier)
                    {
                        // Escaped quote
                        currentValue.Append(c);
                        i++; // Skip next quote
                    }
                    else
                    {
                        // Toggle quote state
                        inQuotes = !inQuotes;
                    }
                }
                else if (c.ToString() == delimiter && !inQuotes)
                {
                    // End of field
                    values.Add(currentValue.ToString());
                    currentValue.Clear();
                }
                else
                {
                    currentValue.Append(c);
                }
            }

            // Add the last field
            values.Add(currentValue.ToString());

            return values.ToArray();
        }

        private static Range DetermineExportRange(Worksheet worksheet, string rangeString)
        {
            if (!string.IsNullOrEmpty(rangeString))
            {
                // Parse range like "A1:E10"
                var parts = rangeString.Split(':');
                if (parts.Length == 2)
                {
                    return worksheet.GetRange(parts[0], parts[1]);
                }
            }

            // If no range specified, find used range
            // For now, return a default range - you might want to implement GetUsedRange()
            return worksheet.GetRange(1, 1, 100, 26); // A1:Z100 as default
        }

        #endregion

        #region JSON Helper Methods

        private static Dictionary<string, object> ProcessJsonObject(JObject jsonObject, JsonOptions options)
        {
            var result = new Dictionary<string, object>();

            foreach (var property in jsonObject.Properties())
            {
                ProcessJsonProperty(result, property.Name, property.Value, options);
            }

            return result;
        }

        private static void ProcessJsonProperty(Dictionary<string, object> result, string propertyName, JToken value, JsonOptions options, string prefix = "")
        {
            var fullPropertyName = string.IsNullOrEmpty(prefix) ? propertyName : $"{prefix}{options.NestedPropertySeparator}{propertyName}";

            switch (value.Type)
            {
                case JTokenType.Object:
                    if (options.FlattenNestedObjects)
                    {
                        var nestedObject = value as JObject;
                        foreach (var nestedProperty in nestedObject.Properties())
                        {
                            ProcessJsonProperty(result, nestedProperty.Name, nestedProperty.Value, options, fullPropertyName);
                        }
                    }
                    else
                    {
                        result[fullPropertyName] = value.ToString();
                    }
                    break;

                case JTokenType.Array:
                    if (options.ConvertArraysToStrings)
                    {
                        var arrayValues = value.Select(v => v.ToString()).ToArray();
                        result[fullPropertyName] = string.Join(options.ArrayValueSeparator, arrayValues);
                    }
                    else
                    {
                        result[fullPropertyName] = value.ToString();
                    }
                    break;

                case JTokenType.Null:
                    if (!options.IgnoreNullValues)
                    {
                        result[fullPropertyName] = null;
                    }
                    break;

                default:
                    result[fullPropertyName] = value.Value<object>();
                    break;
            }
        }

        private static void SetCellValueFromJson(Cell cell, object value, JsonOptions options, CultureInfo culture)
        {
            if (value == null)
            {
                cell.PutValue("");
                return;
            }

            if (!options.AutoDetectDataTypes)
            {
                cell.PutValue(value.ToString());
                return;
            }

            // Try to detect and set appropriate data type
            switch (value)
            {
                case DateTime dateTime:
                    cell.PutValue(dateTime);
                    break;

                case bool boolean:
                    cell.PutValue(boolean.ToString());
                    break;

                case int integer:
                    cell.PutValue(integer);
                    break;

                case long longValue:
                    cell.PutValue((double)longValue);
                    break;

                case float floatValue:
                    cell.PutValue((double)floatValue);
                    break;

                case double doubleValue:
                    cell.PutValue(doubleValue);
                    break;

                case decimal decimalValue:
                    cell.PutValue((double)decimalValue);
                    break;

                default:
                    var stringValue = value.ToString();

                    // Try to parse as DateTime
                    if (DateTime.TryParseExact(stringValue, options.DateTimeFormat, culture, DateTimeStyles.None, out DateTime dateTimeValue))
                    {
                        cell.PutValue(dateTimeValue);
                    }
                    else if (DateTime.TryParseExact(stringValue, options.DateFormat, culture, DateTimeStyles.None, out DateTime dateValue))
                    {
                        cell.PutValue(dateValue);
                    }
                    // Try to parse as number
                    else if (double.TryParse(stringValue, NumberStyles.Any, culture, out double numericValue))
                    {
                        cell.PutValue(numericValue);
                    }
                    // Default to string
                    else
                    {
                        cell.PutValue(stringValue);
                    }
                    break;
            }
        }

        #endregion
    }
}