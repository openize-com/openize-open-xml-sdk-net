using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Openize.Cells
{
    /// <summary>
    /// Manages the application of conditional formatting rules to worksheets.
    /// </summary>
    internal class ConditionalFormattingManager
    {
        private readonly WorksheetPart _worksheetPart;
        private readonly WorkbookStylesPart _stylesPart;
        private readonly List<uint> _differentialFormatIds;

        /// <summary>
        /// Initializes a new instance of the ConditionalFormattingManager class.
        /// </summary>
        public ConditionalFormattingManager(WorksheetPart worksheetPart, WorkbookStylesPart stylesPart)
        {
            _worksheetPart = worksheetPart ?? throw new ArgumentNullException(nameof(worksheetPart));
            _stylesPart = stylesPart ?? throw new ArgumentNullException(nameof(stylesPart));
            _differentialFormatIds = new List<uint>();

            EnsureStylesheetStructure();
        }

        /// <summary>
        /// Applies a conditional formatting rule to the worksheet.
        /// </summary>
        public void ApplyRule(ConditionalFormattingRule rule)
        {
            if (rule == null)
                throw new ArgumentNullException(nameof(rule));

            if (string.IsNullOrEmpty(rule.CellRange))
                throw new ArgumentException("Cell range must be specified for the conditional formatting rule.");

            var conditionalFormatting = CreateConditionalFormatting(rule);
            InsertConditionalFormatting(conditionalFormatting);
        }

        /// <summary>
        /// Applies multiple conditional formatting rules to the worksheet.
        /// </summary>
        public void ApplyRules(IEnumerable<ConditionalFormattingRule> rules)
        {
            foreach (var rule in rules)
            {
                ApplyRule(rule);
            }
        }

        private void EnsureStylesheetStructure()
        {
            if (_stylesPart.Stylesheet == null)
            {
                _stylesPart.Stylesheet = new Stylesheet();
            }

            var stylesheet = _stylesPart.Stylesheet;

            // Ensure DifferentialFormats exists
            if (stylesheet.DifferentialFormats == null)
            {
                stylesheet.DifferentialFormats = new DifferentialFormats() { Count = 0 };
            }

            // Ensure other required elements exist
            if (stylesheet.Fonts == null)
                stylesheet.Fonts = new Fonts(new Font());
            if (stylesheet.Fills == null)
                stylesheet.Fills = new Fills(new Fill());
            if (stylesheet.Borders == null)
                stylesheet.Borders = new Borders(new Border());
            if (stylesheet.CellFormats == null)
                stylesheet.CellFormats = new CellFormats(new CellFormat());
        }

        private ConditionalFormatting CreateConditionalFormatting(ConditionalFormattingRule rule)
        {
            var cf = new ConditionalFormatting
            {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = rule.CellRange }
            };

            var cfRule = CreateConditionalFormattingRule(rule);
            cf.Append(cfRule);

            return cf;
        }

        private DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule CreateConditionalFormattingRule(ConditionalFormattingRule rule)
        {
            var cfRule = new DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule
            {
                Priority = Int32Value.FromInt32((int)rule.Priority),
                StopIfTrue = rule.StopIfTrue
            };

            switch (rule.Type)
            {
                case ConditionalFormattingType.CellIs:
                    ConfigureCellIsRule(cfRule, rule);
                    break;
                case ConditionalFormattingType.DataBar:
                    ConfigureDataBarRule(cfRule, rule);
                    break;
                case ConditionalFormattingType.ColorScale:
                    ConfigureColorScaleRule(cfRule, rule);
                    break;
                case ConditionalFormattingType.DuplicateValues:
                    ConfigureDuplicateValuesRule(cfRule, rule);
                    break;
                case ConditionalFormattingType.ContainsText:
                    ConfigureContainsTextRule(cfRule, rule);
                    break;
                case ConditionalFormattingType.Top10:
                    ConfigureTop10Rule(cfRule, rule);
                    break;
                case ConditionalFormattingType.AboveAverage:
                    ConfigureAboveAverageRule(cfRule, rule);
                    break;
                case ConditionalFormattingType.Expression:
                    ConfigureExpressionRule(cfRule, rule);
                    break;
                default:
                    throw new NotSupportedException($"Conditional formatting type '{rule.Type}' is not supported.");
            }

            return cfRule;
        }

        private void ConfigureCellIsRule(DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule cfRule, ConditionalFormattingRule rule)
        {
            cfRule.Type = ConditionalFormatValues.CellIs;
            cfRule.Operator = ConvertOperator(rule.Operator.Value);

            if (rule.Style != null)
            {
                cfRule.FormatId = CreateDifferentialFormat(rule.Style);
            }

            cfRule.Append(new Formula() { Text = rule.Value1 });

            if (!string.IsNullOrEmpty(rule.Value2) &&
                (rule.Operator == ConditionalFormattingOperator.Between ||
                 rule.Operator == ConditionalFormattingOperator.NotBetween))
            {
                cfRule.Append(new Formula() { Text = rule.Value2 });
            }
        }

        private void ConfigureDataBarRule(DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule cfRule, ConditionalFormattingRule rule)
        {
            cfRule.Type = ConditionalFormatValues.DataBar;

            var dataBar = new DataBar
            {
                ShowValue = rule.DataBar.ShowValue,
                MinLength = (uint)rule.DataBar.MinLength,
                MaxLength = (uint)rule.DataBar.MaxLength
            };

            var cfvo1 = new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min };
            var cfvo2 = new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max };
            var color = new Color { Rgb = rule.DataBar.BarColor };

            dataBar.Append(cfvo1);
            dataBar.Append(cfvo2);
            dataBar.Append(color);

            cfRule.Append(dataBar);
        }

        private void ConfigureColorScaleRule(DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule cfRule, ConditionalFormattingRule rule)
        {
            cfRule.Type = ConditionalFormatValues.ColorScale;

            var colorScale = new ColorScale();

            // Minimum
            var cfvo1 = new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min };
            var color1 = new Color { Rgb = rule.ColorScale.MinColor };
            colorScale.Append(cfvo1);

            // Midpoint (if 3-color scale)
            if (!rule.ColorScale.TwoColorScale)
            {
                var cfvo2 = new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percentile, Val = "50" };
                colorScale.Append(cfvo2);
            }

            // Maximum
            var cfvo3 = new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max };
            colorScale.Append(cfvo3);

            // Add colors
            colorScale.Append(color1);
            if (!rule.ColorScale.TwoColorScale)
            {
                var color2 = new Color { Rgb = rule.ColorScale.MidColor };
                colorScale.Append(color2);
            }
            var color3 = new Color { Rgb = rule.ColorScale.MaxColor };
            colorScale.Append(color3);

            cfRule.Append(colorScale);
        }

        private void ConfigureDuplicateValuesRule(DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule cfRule, ConditionalFormattingRule rule)
        {
            cfRule.Type = ConditionalFormatValues.DuplicateValues;

            if (rule.Style != null)
            {
                cfRule.FormatId = CreateDifferentialFormat(rule.Style);
            }
        }

        private void ConfigureContainsTextRule(DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule cfRule, ConditionalFormattingRule rule)
        {
            cfRule.Type = ConditionalFormatValues.ContainsText;
            cfRule.Operator = ConditionalFormattingOperatorValues.ContainsText;
            cfRule.Text = rule.Value1;

            if (rule.Style != null)
            {
                cfRule.FormatId = CreateDifferentialFormat(rule.Style);
            }

            // Formula for contains text
            string cellRef = rule.CellRange.Split(':')[0];
            cfRule.Append(new Formula() { Text = $"NOT(ISERROR(SEARCH(\"{rule.Value1}\",{cellRef})))" });
        }

        private void ConfigureTop10Rule(DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule cfRule, ConditionalFormattingRule rule)
        {
            cfRule.Type = ConditionalFormatValues.Top10;

            var parts = rule.Value2.Split(',');
            bool isPercent = bool.Parse(parts[0]);
            bool isBottom = bool.Parse(parts[1]);

            cfRule.Rank = uint.Parse(rule.Value1);
            cfRule.Percent = isPercent;
            cfRule.Bottom = isBottom;

            if (rule.Style != null)
            {
                cfRule.FormatId = CreateDifferentialFormat(rule.Style);
            }
        }

        private void ConfigureAboveAverageRule(DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule cfRule, ConditionalFormattingRule rule)
        {
            cfRule.Type = ConditionalFormatValues.AboveAverage;
            bool isAbove = bool.Parse(rule.Value1);
            cfRule.AboveAverage = isAbove;

            if (rule.Style != null)
            {
                cfRule.FormatId = CreateDifferentialFormat(rule.Style);
            }
        }

        private void ConfigureExpressionRule(DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule cfRule, ConditionalFormattingRule rule)
        {
            cfRule.Type = ConditionalFormatValues.Expression;

            if (rule.Style != null)
            {
                cfRule.FormatId = CreateDifferentialFormat(rule.Style);
            }

            cfRule.Append(new Formula() { Text = rule.Value1 });
        }

        private uint CreateDifferentialFormat(ConditionalFormattingStyle style)
        {
            var df = new DifferentialFormat();

            // Font
            if (!string.IsNullOrEmpty(style.FontColor) || style.IsBold.HasValue || style.IsItalic.HasValue)
            {
                var font = new Font();

                if (!string.IsNullOrEmpty(style.FontColor))
                    font.Append(new Color { Rgb = style.FontColor });

                if (style.IsBold.HasValue && style.IsBold.Value)
                    font.Append(new Bold());

                if (style.IsItalic.HasValue && style.IsItalic.Value)
                    font.Append(new Italic());

                df.Append(font);
            }

            // Fill
            if (!string.IsNullOrEmpty(style.BackgroundColor))
            {
                var fill = new Fill();
                var patternFill = new PatternFill { PatternType = style.PatternType ?? PatternValues.Solid };

                if (style.PatternType == PatternValues.Solid)
                {
                    patternFill.Append(new ForegroundColor { Rgb = style.BackgroundColor });
                    patternFill.Append(new BackgroundColor { Rgb = style.BackgroundColor });
                }
                else
                {
                    patternFill.Append(new BackgroundColor { Rgb = style.BackgroundColor });
                }

                fill.Append(patternFill);
                df.Append(fill);
            }

            // Border
            if (!string.IsNullOrEmpty(style.BorderColor))
            {
                var border = new Border();
                var color = new Color { Rgb = style.BorderColor };

                border.Append(new LeftBorder { Style = BorderStyleValues.Thin, Color = color });
                border.Append(new RightBorder { Style = BorderStyleValues.Thin, Color = color });
                border.Append(new TopBorder { Style = BorderStyleValues.Thin, Color = color });
                border.Append(new BottomBorder { Style = BorderStyleValues.Thin, Color = color });

                df.Append(border);
            }

            // Add to stylesheet
            var dfs = _stylesPart.Stylesheet.DifferentialFormats;
            dfs.Append(df);
            dfs.Count = (uint)dfs.Count();

            _stylesPart.Stylesheet.Save();

            return (uint)(dfs.Count() - 1);
        }

        private ConditionalFormattingOperatorValues ConvertOperator(ConditionalFormattingOperator op)
        {
            return op switch
            {
                ConditionalFormattingOperator.LessThan => ConditionalFormattingOperatorValues.LessThan,
                ConditionalFormattingOperator.LessThanOrEqual => ConditionalFormattingOperatorValues.LessThanOrEqual,
                ConditionalFormattingOperator.Equal => ConditionalFormattingOperatorValues.Equal,
                ConditionalFormattingOperator.NotEqual => ConditionalFormattingOperatorValues.NotEqual,
                ConditionalFormattingOperator.GreaterThanOrEqual => ConditionalFormattingOperatorValues.GreaterThanOrEqual,
                ConditionalFormattingOperator.GreaterThan => ConditionalFormattingOperatorValues.GreaterThan,
                ConditionalFormattingOperator.Between => ConditionalFormattingOperatorValues.Between,
                ConditionalFormattingOperator.NotBetween => ConditionalFormattingOperatorValues.NotBetween,
                ConditionalFormattingOperator.Contains => ConditionalFormattingOperatorValues.ContainsText,
                ConditionalFormattingOperator.NotContains => ConditionalFormattingOperatorValues.NotContains,
                ConditionalFormattingOperator.BeginsWith => ConditionalFormattingOperatorValues.BeginsWith,
                ConditionalFormattingOperator.EndsWith => ConditionalFormattingOperatorValues.EndsWith,
                _ => throw new ArgumentException($"Unsupported operator: {op}")
            };
        }

        private void InsertConditionalFormatting(ConditionalFormatting cf)
        {
            var worksheet = _worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();

            // Insert after SheetData but before other elements like MergeCells
            var insertAfter = sheetData;
            var existingCfs = worksheet.Elements<ConditionalFormatting>();
            if (existingCfs.Any())
            {
                worksheet.InsertAfter<ConditionalFormatting>(cf, insertAfter);
            }

            worksheet.InsertAfter(cf.CloneNode(true), insertAfter);
        }
    }
}