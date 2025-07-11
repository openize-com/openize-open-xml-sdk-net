using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Openize.Cells
{
    /// <summary>
    /// Represents the type of conditional formatting rule.
    /// </summary>
    public enum ConditionalFormattingType
    {
        /// <summary>Cell value based formatting</summary>
        CellIs,
        /// <summary>Data bar visualization</summary>
        DataBar,
        /// <summary>Color scale visualization</summary>
        ColorScale,
        /// <summary>Icon set visualization</summary>
        IconSet,
        /// <summary>Top/Bottom rules</summary>
        Top10,
        /// <summary>Above/Below average</summary>
        AboveAverage,
        /// <summary>Duplicate values</summary>
        DuplicateValues,
        /// <summary>Contains text</summary>
        ContainsText,
        /// <summary>Date occurring</summary>
        TimePeriod,
        /// <summary>Expression based</summary>
        Expression
    }

    /// <summary>
    /// Represents the operator for conditional formatting rules.
    /// </summary>
    public enum ConditionalFormattingOperator
    {
        /// <summary>Less than</summary>
        LessThan,
        /// <summary>Less than or equal</summary>
        LessThanOrEqual,
        /// <summary>Equal</summary>
        Equal,
        /// <summary>Not equal</summary>
        NotEqual,
        /// <summary>Greater than or equal</summary>
        GreaterThanOrEqual,
        /// <summary>Greater than</summary>
        GreaterThan,
        /// <summary>Between</summary>
        Between,
        /// <summary>Not between</summary>
        NotBetween,
        /// <summary>Contains</summary>
        Contains,
        /// <summary>Not contains</summary>
        NotContains,
        /// <summary>Begins with</summary>
        BeginsWith,
        /// <summary>Ends with</summary>
        EndsWith
    }

    /// <summary>
    /// Represents a conditional formatting style definition.
    /// </summary>
    public class ConditionalFormattingStyle
    {
        /// <summary>Gets or sets the background color in hex format (e.g., "FF92D050" for green)</summary>
        public string BackgroundColor { get; set; }

        /// <summary>Gets or sets the font color in hex format (e.g., "FFFFFFFF" for white)</summary>
        public string FontColor { get; set; }

        /// <summary>Gets or sets whether the font is bold</summary>
        public bool? IsBold { get; set; }

        /// <summary>Gets or sets whether the font is italic</summary>
        public bool? IsItalic { get; set; }

        /// <summary>Gets or sets the border color in hex format</summary>
        public string BorderColor { get; set; }

        /// <summary>Gets or sets the pattern type for fill</summary>
        public PatternValues? PatternType { get; set; } = PatternValues.Solid;
    }

    /// <summary>
    /// Represents a data bar configuration for conditional formatting.
    /// </summary>
    public class DataBarOptions
    {
        /// <summary>Gets or sets the color of the data bar in hex format</summary>
        public string BarColor { get; set; } = "FF638EC6"; // Default blue

        /// <summary>Gets or sets whether to show the cell value</summary>
        public bool ShowValue { get; set; } = true;

        /// <summary>Gets or sets the minimum length percentage (0-100)</summary>
        public int MinLength { get; set; } = 10;

        /// <summary>Gets or sets the maximum length percentage (0-100)</summary>
        public int MaxLength { get; set; } = 90;
    }

    /// <summary>
    /// Represents a color scale configuration for conditional formatting.
    /// </summary>
    public class ColorScaleOptions
    {
        /// <summary>Gets or sets the minimum value color in hex format</summary>
        public string MinColor { get; set; } = "FFF8696B"; // Red

        /// <summary>Gets or sets the midpoint value color in hex format</summary>
        public string MidColor { get; set; } = "FFFFEB84"; // Yellow

        /// <summary>Gets or sets the maximum value color in hex format</summary>
        public string MaxColor { get; set; } = "FF63BE7B"; // Green

        /// <summary>Gets or sets whether to use a 2-color scale (if false, uses 3-color scale)</summary>
        public bool TwoColorScale { get; set; } = false;
    }

    /// <summary>
    /// Represents a conditional formatting rule.
    /// </summary>
    public class ConditionalFormattingRule
    {
        private static uint _nextPriority = 1;

        /// <summary>Gets or sets the type of conditional formatting</summary>
        public ConditionalFormattingType Type { get; set; }

        /// <summary>Gets or sets the operator for comparison</summary>
        public ConditionalFormattingOperator? Operator { get; set; }

        /// <summary>Gets or sets the first value or formula for comparison</summary>
        public string Value1 { get; set; }

        /// <summary>Gets or sets the second value for between/not between operators</summary>
        public string Value2 { get; set; }

        /// <summary>Gets or sets the style to apply when condition is met</summary>
        public ConditionalFormattingStyle Style { get; set; }

        /// <summary>Gets or sets the data bar options</summary>
        public DataBarOptions DataBar { get; set; }

        /// <summary>Gets or sets the color scale options</summary>
        public ColorScaleOptions ColorScale { get; set; }

        /// <summary>Gets or sets the priority of the rule</summary>
        public uint Priority { get; private set; }

        /// <summary>Gets or sets whether to stop if this rule is true</summary>
        public bool StopIfTrue { get; set; } = false;

        /// <summary>Gets or sets the cell range(s) to apply the formatting to</summary>
        public string CellRange { get; set; }

        /// <summary>
        /// Initializes a new instance of the ConditionalFormattingRule class.
        /// </summary>
        public ConditionalFormattingRule()
        {
            Priority = _nextPriority++;
        }

        /// <summary>
        /// Creates a rule for highlighting cells greater than a value.
        /// </summary>
        public static ConditionalFormattingRule GreaterThan(string value, ConditionalFormattingStyle style)
        {
            return new ConditionalFormattingRule
            {
                Type = ConditionalFormattingType.CellIs,
                Operator = ConditionalFormattingOperator.GreaterThan,
                Value1 = value,
                Style = style
            };
        }

        /// <summary>
        /// Creates a rule for highlighting cells less than a value.
        /// </summary>
        public static ConditionalFormattingRule LessThan(string value, ConditionalFormattingStyle style)
        {
            return new ConditionalFormattingRule
            {
                Type = ConditionalFormattingType.CellIs,
                Operator = ConditionalFormattingOperator.LessThan,
                Value1 = value,
                Style = style
            };
        }

        /// <summary>
        /// Creates a rule for highlighting cells between two values.
        /// </summary>
        public static ConditionalFormattingRule Between(string value1, string value2, ConditionalFormattingStyle style)
        {
            return new ConditionalFormattingRule
            {
                Type = ConditionalFormattingType.CellIs,
                Operator = ConditionalFormattingOperator.Between,
                Value1 = value1,
                Value2 = value2,
                Style = style
            };
        }

        /// <summary>
        /// Creates a rule for highlighting cells equal to a value.
        /// </summary>
        public static ConditionalFormattingRule EqualTo(string value, ConditionalFormattingStyle style)
        {
            return new ConditionalFormattingRule
            {
                Type = ConditionalFormattingType.CellIs,
                Operator = ConditionalFormattingOperator.Equal,
                Value1 = value,
                Style = style
            };
        }

        /// <summary>
        /// Creates a rule for data bar visualization.
        /// </summary>
        public static ConditionalFormattingRule DataBarRule(DataBarOptions options = null)
        {
            return new ConditionalFormattingRule
            {
                Type = ConditionalFormattingType.DataBar,
                DataBar = options ?? new DataBarOptions()
            };
        }

        /// <summary>
        /// Creates a rule for color scale visualization.
        /// </summary>
        public static ConditionalFormattingRule ColorScaleRule(ColorScaleOptions options = null)
        {
            return new ConditionalFormattingRule
            {
                Type = ConditionalFormattingType.ColorScale,
                ColorScale = options ?? new ColorScaleOptions()
            };
        }

        /// <summary>
        /// Creates a rule for highlighting duplicate values.
        /// </summary>
        public static ConditionalFormattingRule DuplicateValues(ConditionalFormattingStyle style)
        {
            return new ConditionalFormattingRule
            {
                Type = ConditionalFormattingType.DuplicateValues,
                Style = style
            };
        }

        /// <summary>
        /// Creates a rule for highlighting cells containing specific text.
        /// </summary>
        public static ConditionalFormattingRule ContainsText(string text, ConditionalFormattingStyle style)
        {
            return new ConditionalFormattingRule
            {
                Type = ConditionalFormattingType.ContainsText,
                Operator = ConditionalFormattingOperator.Contains,
                Value1 = text,
                Style = style
            };
        }

        /// <summary>
        /// Creates a rule for highlighting top N values.
        /// </summary>
        public static ConditionalFormattingRule Top10(int count, bool isPercent, bool isBottom, ConditionalFormattingStyle style)
        {
            return new ConditionalFormattingRule
            {
                Type = ConditionalFormattingType.Top10,
                Value1 = count.ToString(),
                Value2 = $"{isPercent},{isBottom}",
                Style = style
            };
        }

        /// <summary>
        /// Creates a rule for highlighting above average values.
        /// </summary>
        public static ConditionalFormattingRule AboveAverage(bool isAbove, ConditionalFormattingStyle style)
        {
            return new ConditionalFormattingRule
            {
                Type = ConditionalFormattingType.AboveAverage,
                Value1 = isAbove.ToString(),
                Style = style
            };
        }

        /// <summary>
        /// Creates a rule based on a custom formula.
        /// </summary>
        public static ConditionalFormattingRule Formula(string formula, ConditionalFormattingStyle style)
        {
            return new ConditionalFormattingRule
            {
                Type = ConditionalFormattingType.Expression,
                Value1 = formula,
                Style = style
            };
        }
    }
}