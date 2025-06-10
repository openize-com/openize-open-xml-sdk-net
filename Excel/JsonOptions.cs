using System;

namespace Openize.Cells
{
    /// <summary>
    /// Configuration options for JSON import and export operations.
    /// </summary>
    public class JsonOptions
    {
        /// <summary>
        /// Gets or sets whether to include property names as headers in the first row.
        /// </summary>
        public bool IncludeHeaders { get; set; } = true;

        /// <summary>
        /// Gets or sets the date format for parsing and formatting date values.
        /// </summary>
        public string DateFormat { get; set; } = "yyyy-MM-dd";

        /// <summary>
        /// Gets or sets the date-time format for parsing and formatting datetime values.
        /// </summary>
        public string DateTimeFormat { get; set; } = "yyyy-MM-dd HH:mm:ss";

        /// <summary>
        /// Gets or sets whether to auto-detect data types during import.
        /// </summary>
        public bool AutoDetectDataTypes { get; set; } = true;

        /// <summary>
        /// Gets or sets the number format culture (e.g., "en-US", "de-DE").
        /// </summary>
        public string Culture { get; set; } = "en-US";

        /// <summary>
        /// Gets or sets whether to ignore null or empty values during import.
        /// </summary>
        public bool IgnoreNullValues { get; set; } = false;

        /// <summary>
        /// Gets or sets the maximum number of records to import (0 = no limit).
        /// </summary>
        public int MaxRecords { get; set; } = 0;

        /// <summary>
        /// Gets or sets whether to flatten nested objects into columns with dot notation.
        /// Example: {"user": {"name": "John"}} becomes "user.name" column.
        /// </summary>
        public bool FlattenNestedObjects { get; set; } = false;

        /// <summary>
        /// Gets or sets the separator for flattened nested object property names.
        /// </summary>
        public string NestedPropertySeparator { get; set; } = ".";

        /// <summary>
        /// Gets or sets whether to handle arrays by converting them to comma-separated strings.
        /// </summary>
        public bool ConvertArraysToStrings { get; set; } = true;

        /// <summary>
        /// Gets or sets the separator for array values when converting to strings.
        /// </summary>
        public string ArrayValueSeparator { get; set; } = ", ";
    }
}