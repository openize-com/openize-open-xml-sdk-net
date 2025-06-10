using System;
using System.Text;

namespace Openize.Cells
{
    /// <summary>
    /// Configuration options for CSV import and export operations.
    /// </summary>
    public class CsvOptions
    {
        /// <summary>
        /// Gets or sets whether the first row contains headers.
        /// </summary>
        public bool HasHeaders { get; set; } = true;

        /// <summary>
        /// Gets or sets the delimiter character used to separate values.
        /// </summary>
        public string Delimiter { get; set; } = ",";

        /// <summary>
        /// Gets or sets the text qualifier character (e.g., quote character).
        /// </summary>
        public string TextQualifier { get; set; } = "\"";

        /// <summary>
        /// Gets or sets the date format for parsing date values.
        /// </summary>
        public string DateFormat { get; set; } = "yyyy-MM-dd";

        /// <summary>
        /// Gets or sets the number format culture (e.g., "en-US", "de-DE").
        /// </summary>
        public string Culture { get; set; } = "en-US";

        /// <summary>
        /// Gets or sets whether to skip empty lines during import.
        /// </summary>
        public bool SkipEmptyLines { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to trim whitespace from values.
        /// </summary>
        public bool TrimWhitespace { get; set; } = true;

        /// <summary>
        /// Gets or sets the encoding for reading/writing the CSV file.
        /// </summary>
        public Encoding Encoding { get; set; } = Encoding.UTF8;

        /// <summary>
        /// Gets or sets the maximum number of rows to import (0 = no limit).
        /// </summary>
        public int MaxRows { get; set; } = 0;

        /// <summary>
        /// Gets or sets whether to auto-detect data types.
        /// </summary>
        public bool AutoDetectDataTypes { get; set; } = true;
    }
}