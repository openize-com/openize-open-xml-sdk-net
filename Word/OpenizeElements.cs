﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Openize.Words.IElements
{
    /// <summary>
    /// Represents an element in a Word document.
    /// </summary>
    public interface IElement
    {
        /// <summary>
        /// Gets the unique identifier of the element.
        /// </summary>
        int ElementId { get; }
    }

    /// <summary>
    /// Represents the indentation settings for a paragraph.
    /// </summary>
    public class Indentation
    {
        /// <summary>
        /// Gets or sets the distance of the left indentation.
        /// </summary>
        public double Left { get; set; }

        /// <summary>
        /// Gets or sets the distance of the right indentation.
        /// </summary>
        public double Right { get; set; }

        /// <summary>
        /// Gets or sets the distance of the first line indentation.
        /// </summary>
        public double FirstLine { get; set; }

        /// <summary>
        /// Gets or sets the distance of the hanging indentation.
        /// </summary>
        public double Hanging { get; set; }

        //public Indentation()
        //{
        //Left = 0;
        //}
    }

    /// <summary>
    /// Specifies the alignment of a paragraph within a text block or document.
    /// </summary>
    public enum ParagraphAlignment
    {
        /// <summary>
        /// Aligns the paragraph to the left.
        /// </summary>
        Left,

        /// <summary>
        /// Centers the paragraph within the available space.
        /// </summary>
        Center,

        /// <summary>
        /// Aligns the paragraph to the right.
        /// </summary>
        Right,

        /// <summary>
        /// Justifies the text within the paragraph, aligning both the left and right edges.
        /// </summary>
        Justify
    }

    /// <summary>
    /// Specifies the border width of an element within a text block or document.
    /// </summary>
    public enum BorderWidth
    {
        /// <summary>
        /// Single width border/frame.
        /// </summary>
        Single,

        /// <summary>
        /// Double width border/frame.
        /// </summary>
        Double,

        /// <summary>
        /// Dotted style border/frame.
        /// </summary>
        Dotted,

        /// <summary>
        /// DotDash style border/frame.
        /// </summary>
        DotDash
    }

    /// <summary>
    /// Represents border/frame of an element within word document.
    /// </summary>
    public class Border
    {
        /// <summary>
        /// Gets or sets the border width.
        /// </summary>
        public BorderWidth Width { get; set; }

        /// <summary>
        /// Gets or sets the border color.
        /// </summary>
        public string Color { get; set; }

        /// <summary>
        /// Gets or sets the border color.
        /// </summary>
        public int Size { get; set; }

        /// <summary>
        /// Constructor for border.
        /// </summary>
        public Border()
        {
            Size = 0;
        }
    }

    /// <summary>
    /// Represents a paragraph element in a Word document.
    /// </summary>
    public class Paragraph : IElement
    {
        /// <summary>
        /// Gets the unique identifier of the paragraph.
        /// </summary>
        public int ElementId { get; internal set; }

        /// <summary>
        /// Gets or sets the text content of the paragraph.
        /// </summary>
        public string Text { get; private set; }

        /// <summary>
        /// Gets the list of runs (text fragments) within the paragraph.
        /// </summary>
        public List<Run> Runs { get; }

        /// <summary>
        /// Gets or sets the style of the paragraph.
        /// </summary>
        public string Style { get; set; }

        /// <summary>
        /// Gets or Sets Alignment of the word paragraph
        /// </summary>
        public ParagraphAlignment Alignment { get; set; }

        /// <summary>
        /// Gets or Sets Indentation of the word paragraph
        /// </summary>
        public Indentation Indentation { get; set; }

        /// <summary>
        /// Gets or sets the numbering ID for the paragraph.
        /// </summary>
        public int? NumberingId { get; set; }

        /// <summary>
        /// Gets or sets the numbering level for the paragraph.
        /// </summary>
        public int? NumberingLevel { get; set; }

        /// <summary>
        /// Gets or sets whether the paragraph has bullet points.
        /// </summary>
        public bool IsBullet { get; set; }

        /// <summary>
        /// Gets or sets whether the paragraph has numbered bullets.
        /// </summary>
        public bool IsNumbered { get; set; }

        /// <summary>
        /// Gets or sets whether the paragraph has roman number bullets.
        /// </summary>
        public bool IsRoman { get; set; }

        /// <summary>
        /// Gets or sets whether the paragraph has alphabetic number bullets.
        /// </summary>
        public bool IsAlphabeticNumber { get; set; }

        /// <summary>
        /// Gets or sets the paragraph border.
        /// </summary>
        public Border ParagraphBorder { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Paragraph"/> class.
        /// </summary>
        public Paragraph()
        {
            Runs = new List<Run>();
            Style = "Normal";
            ParagraphBorder = new Border();
            Indentation = new Indentation();
            NumberingId = null;
            NumberingLevel = null;
            IsBullet = false;
            IsNumbered = false;
            UpdateText(); // Initialize the Text property
        }

        /// <summary>
        /// Adds a run (text fragment) to the paragraph and sets its parent paragraph.
        /// </summary>
        /// <param name="run">The run to add to the paragraph.</param>
        public void AddRun(Run run)
        {
            run.ParentParagraph = this;
            Runs.Add(run);
            UpdateText(); // Update the Text property when a new run is added
        }

        internal void UpdateText()
        {
            Text = string.Join("", Runs.Select(run => run.Text));
        }

        internal void ReplaceText(string search, string replacement, bool useRegex = false)
        {
            var matches = useRegex
            ? Regex.Matches(this.Text, search)
            : Regex.Matches(this.Text, Regex.Escape(search));

            if (matches.Count == 0) return;

            var newRuns = new List<Run>();
            int globalIndex = 0;

            foreach (Match match in matches)
            {
                int matchStart = match.Index;
                int matchEnd = match.Index + match.Length;

                // Track through original runs
                int currentIndex = 0;
                foreach (var run in this.Runs)
                {
                    var text = run.Text;
                    int runStart = currentIndex;
                    int runEnd = currentIndex + text.Length;

                    if (runEnd <= matchStart || runStart >= matchEnd)
                    {
                        newRuns.Add(CloneRun(run));
                    }
                    else
                    {
                        // Overlaps with the match
                        int relativeStart = Math.Max(0, matchStart - currentIndex);
                        int relativeEnd = Math.Min(text.Length, matchEnd - currentIndex);

                        if (relativeStart > 0)
                        {
                            newRuns.Add(new Run
                            {
                                Text = text.Substring(0, relativeStart),
                                FontFamily = run.FontFamily,
                                FontSize = run.FontSize,
                                Color = run.Color,
                                Bold = run.Bold,
                                Italic = run.Italic,
                                Underline = run.Underline
                            });
                        }

                        if (runStart <= matchStart && runEnd >= matchEnd)
                        {
                            newRuns.Add(new Run
                            {
                                Text = replacement,
                                FontFamily = run.FontFamily,
                                FontSize = run.FontSize,
                                Color = run.Color,
                                Bold = run.Bold,
                                Italic = run.Italic,
                                Underline = run.Underline
                            });
                        }

                        if (relativeEnd < text.Length)
                        {
                            newRuns.Add(new Run
                            {
                                Text = text.Substring(relativeEnd),
                                FontFamily = run.FontFamily,
                                FontSize = run.FontSize,
                                Color = run.Color,
                                Bold = run.Bold,
                                Italic = run.Italic,
                                Underline = run.Underline
                            });
                        }
                    }

                    currentIndex += run.Text.Length;

                    if (currentIndex >= matchEnd)
                        break;
                }

                // Replace only the first match at a time to avoid index shifting
                break;
            }

            this.Runs.Clear();
            foreach (var r in newRuns)
            {
                this.AddRun(r);
            }
        }
        private static Run CloneRun(Run r) =>
        new Run
        {
            Text = r.Text,
            FontFamily = r.FontFamily,
            FontSize = r.FontSize,
            Color = r.Color,
            Bold = r.Bold,
            Italic = r.Italic,
            Underline = r.Underline
        };
    }

    /// <summary>
    /// Provides predefined heading styles.
    /// </summary>
    public static class Headings
    {
        /// <summary>
        /// Gets the value representing Heading1.
        /// </summary>
        public static string Heading1 { get; } = "Heading1";

        /// <summary>
        /// Gets the value representing Heading2.
        /// </summary>
        public static string Heading2 { get; } = "Heading2";

        /// <summary>
        /// Gets the value representing Heading3.
        /// </summary>
        public static string Heading3 { get; } = "Heading3";

        /// <summary>
        /// Gets the value representing Heading4.
        /// </summary>
        public static string Heading4 { get; } = "Heading4";

        /// <summary>
        /// Gets the value representing Heading5.
        /// </summary>
        public static string Heading5 { get; } = "Heading5";

        /// <summary>
        /// Gets the value representing Heading6.
        /// </summary>
        public static string Heading6 { get; } = "Heading6";

        /// <summary>
        /// Gets the value representing Heading7.
        /// </summary>
        public static string Heading7 { get; } = "Heading7";

        /// <summary>
        /// Gets the value representing Heading8.
        /// </summary>
        public static string Heading8 { get; } = "Heading8";

        /// <summary>
        /// Gets the value representing Heading9.
        /// </summary>
        public static string Heading9 { get; } = "Heading9";
    }


    /// <summary>
    /// Represents a run of text within a paragraph.
    /// </summary>
    public class Run
    {
        private string _text;
        /// <summary>
        /// Gets or sets the text content of the run.
        /// </summary>
        public string Text
        {

            get => _text;
            set
            {
                _text = value;
                if (ParentParagraph != null)
                {
                    ParentParagraph.UpdateText();
                }
            }
        }

        /// <summary>
        /// Gets or sets the font family of the run.
        /// </summary>
        public string FontFamily { get; set; }

        /// <summary>
        /// Gets or sets the font size of the run.
        /// </summary>
        public int FontSize { get; set; }

        /// <summary>
        /// Gets or sets the color of the run's text.
        /// </summary>
        public string Color { get; set; }

        /// <summary>
        /// Gets or sets whether the run's text is bold.
        /// </summary>
        public bool Bold { get; set; }

        /// <summary>
        /// Gets or sets whether the run's text is italic.
        /// </summary>
        public bool Italic { get; set; }

        /// <summary>
        /// Gets or sets whether the run's text is underlined.
        /// </summary>
        public bool Underline { get; set; }

        internal Paragraph ParentParagraph { get; set; }
    }

    /// <summary>
    /// Provides predefined colors with hexadecimal values.
    /// </summary>
    public static class Colors
    {
        /// <summary>
        /// Gets the hexadecimal value for the color Black (000000).
        /// </summary>
        public static string Black { get; } = "000000";

        /// <summary>
        /// Gets the hexadecimal value for the color White (FFFFFF).
        /// </summary>
        public static string White { get; } = "FFFFFF";

        /// <summary>
        /// Gets the hexadecimal value for the color Red (FF0000).
        /// </summary>
        public static string Red { get; } = "FF0000";

        /// <summary>
        /// Gets the hexadecimal value for the color Green (00FF00).
        /// </summary>
        public static string Green { get; } = "00FF00";

        /// <summary>
        /// Gets the hexadecimal value for the color Blue (0000FF).
        /// </summary>
        public static string Blue { get; } = "0000FF";

        /// <summary>
        /// Gets the hexadecimal value for the color Yellow (FFFF00).
        /// </summary>
        public static string Yellow { get; } = "FFFF00";

        /// <summary>
        /// Gets the hexadecimal value for the color Cyan (00FFFF).
        /// </summary>
        public static string Cyan { get; } = "00FFFF";

        /// <summary>
        /// Gets the hexadecimal value for the color Magenta (FF00FF).
        /// </summary>
        public static string Magenta { get; } = "FF00FF";

        /// <summary>
        /// Gets the hexadecimal value for the color Gray (808080).
        /// </summary>
        public static string Gray { get; } = "808080";

        /// <summary>
        /// Gets the hexadecimal value for the color Silver (C0C0C0).
        /// </summary>
        public static string Silver { get; } = "C0C0C0";

        /// <summary>
        /// Gets the hexadecimal value for the color Maroon (800000).
        /// </summary>
        public static string Maroon { get; } = "800000";

        /// <summary>
        /// Gets the hexadecimal value for the color Olive (808000).
        /// </summary>
        public static string Olive { get; } = "808000";

        /// <summary>
        /// Gets the hexadecimal value for the color Green (008000).
        /// </summary>
        public static string Teal { get; } = "008000";

        /// <summary>
        /// Gets the hexadecimal value for the color Navy (000080).
        /// </summary>
        public static string Navy { get; } = "000080";

        /// <summary>
        /// Gets the hexadecimal value for the color Purple (800080).
        /// </summary>
        public static string Purple { get; } = "800080";

        /// <summary>
        /// Gets the hexadecimal value for the color Orange (FFA500).
        /// </summary>
        public static string Orange { get; } = "FFA500";

        /// <summary>
        /// Gets the hexadecimal value for the color Lime (00FF00).
        /// </summary>
        public static string Lime { get; } = "00FF00";

        /// <summary>
        /// Gets the hexadecimal value for the color Aqua (00FFFF).
        /// </summary>
        public static string Aqua { get; } = "00FFFF";

        /// <summary>
        /// Gets the hexadecimal value for the color Fuchsia (FF00FF).
        /// </summary>
        public static string Fuchsia { get; } = "FF00FF";

        /// <summary>
        /// Gets the hexadecimal value for the color Silver (C0C0C0).
        /// </summary>
        public static string LimeGreen { get; } = "32CD32";
    }

    /// <summary>
    /// Represents an image element in a Word document.
    /// </summary>
    public class Image : IElement
    {
        /// <summary>
        /// Gets the unique identifier of the image.
        /// </summary>
        public int ElementId { get; internal set; }

        /// <summary>
        /// Gets or sets the binary image data.
        /// </summary>
        public byte[] ImageData { get; set; }

        /// <summary>
        /// Gets or sets the height of the image.
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// Gets or sets the width of the image.
        /// </summary>
        public int Width { get; set; }
    }

    /// <summary>
    /// Represents a shape element in a Word document.
    /// </summary>
    public class Shape : IElement
    {
        /// <summary>
        /// Gets the unique identifier of the shape.
        /// </summary>
        public int ElementId { get; internal set; }

        /// <summary>
        /// Gets or sets the x position of the shape.
        /// </summary>
        public int X { get; set; }

        /// <summary>
        /// Gets or sets the y position of the shape.
        /// </summary>
        public int Y { get; set; }

        /// <summary>
        /// Gets or sets the height of the shape.
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// Gets or sets the width of the shape.
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Gets or sets the type of the shape.
        /// </summary>
        public ShapeType Type { get; set; }

        /// <summary>
        /// Gets or sets the fill type of the shape
        /// </summary>
        public ShapeFillType FillType { get; set; }

        /// <summary>
        /// Gets or sets the fill colors of the shape
        /// </summary>
        public ShapeFillColors FillColors { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Shape"/> class.
        /// </summary>
        public Shape()
        {
            X = 100;
            Y = 100;
            Width = 200;
            Height = 200;
            Type = ShapeType.Ellipse;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Shape"/> class with specified values.
        /// </summary>
        /// <param name="x">x position of the shape.</param>
        /// <param name="y">y position of the shape.</param>
        /// <param name="width">Width of the shape.</param>
        /// <param name="height">Height of the shape.</param>
        /// <param name="shapeType">Type of the shape (e.g rectangle, ellipse etc).</param>
        public Shape(int x, int y, int width, int height, ShapeType shapeType)
        {
            X = x;
            Y = y;
            Width = width;
            Height = height;
            Type = shapeType;
            FillType = ShapeFillType.None;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Shape"/> class with more detailed values.
        /// </summary>
        /// <param name="x">x position of the shape.</param>
        /// <param name="y">y position of the shape.</param>
        /// <param name="width">Width of the shape.</param>
        /// <param name="height">Height of the shape.</param>
        /// <param name="shapeType">Type of the shape (e.g rectangle, ellipse etc).</param>
        /// <param name="shapeFillType">Fill type of the shape (e.g solid, gradient, pattern).</param>
        /// <param name="shapeFillColors">Fill colors of the shape.</param>
        public Shape(int x, int y, int width, int height,
            ShapeType shapeType,
            ShapeFillType shapeFillType,
            ShapeFillColors shapeFillColors)
        {
            X = x;
            Y = y;
            Width = width;
            Height = height;
            Type = shapeType;
            FillType = shapeFillType;
            FillColors = shapeFillColors;
        }
    }

    /// <summary>
    /// Represents a grouped shapes element in a Word document.
    /// </summary>
    public class GroupShape : IElement
    {
        /// <summary>
        /// Gets the unique identifier of the group shape.
        /// </summary>
        public int ElementId { get; internal set; }

        /// <summary>
        /// Gets or sets the type of the first shape in the group.
        /// </summary>
        public Shape Shape1 { get; set; }

        /// <summary>
        /// Gets or sets the type of the second shape in the group.
        /// </summary>
        public Shape Shape2 { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="GroupShape"/> class.
        /// </summary>
        public GroupShape()
        {
            Shape1 = new Shape();
            Shape2 = new Shape();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GroupShape"/> class with specified values.
        /// </summary>
        /// <param name="shape1">Firs shape in the group.</param>
        /// <param name="shape2">Second shape in the group.</param>
        public GroupShape(Shape shape1, Shape shape2)
        {
            Shape1 = shape1;
            //Shape1.ElementId = ElementId + 1;
            Shape2 = shape2;
            //Shape2.ElementId = ElementId + 2;
        }
    }

    /// <summary>
    /// Specifies the type of a shape within the word document.
    /// </summary>
    public enum ShapeType
    {
        /// <summary>
        /// Rectangle shape.
        /// </summary>
        Rectangle,

        /// <summary>
        /// Triangle shape.
        /// </summary>
        Triangle,

        /// <summary>
        /// Ellipse or Oval shape.
        /// </summary>
        Ellipse,

        /// <summary>
        /// Diamond shape.
        /// </summary>
        Diamond,

        /// <summary>
        /// Hexagone shape.
        /// </summary>
        Hexagone
    }

    /// <summary>
    /// Specifies the fill type of a shape within the word document.
    /// </summary>
    public enum ShapeFillType
    {
        /// <summary>
        /// Solid fill.
        /// </summary>
        Solid,

        /// <summary>
        /// Gradient fill.
        /// </summary>
        Gradient,

        /// <summary>
        /// Pattern fill.
        /// </summary>
        Pattern,

        /// <summary>
        /// Default fill.
        /// </summary>
        None
    }

    /// <summary>
    /// Represents shape fill colors.
    /// </summary>
    public class ShapeFillColors
    {
        /// <summary>
        /// Initializes with default colors.
        /// </summary>
        public ShapeFillColors()
        {
            Color1 = Colors.Red;
            Color2 = Colors.Purple;
        }

        public ShapeFillColors(string color1,string color2)
        {
            Color1 = color1;
            Color2 = color2;
        }

        /// <summary>
        /// Gets or sets the first color in gradient/pattern fill and/or color in solid fill
        /// </summary>
        public string Color1 { get; set; }

        /// <summary>
        /// Gets or sets the second color in gradient/pattern fill
        /// </summary>
        public string Color2 { get; set; }
    }

    /// <summary>
    /// Represents a table element in a Word document.
    /// </summary>
    public class Table : IElement
    {
        /// <summary>
        /// Gets the unique identifier of the table.
        /// </summary>
        public int ElementId { get; internal set; }

        /// <summary>
        /// Gets or sets the table style.
        /// </summary>
        public string Style { get; set; }

        /// <summary>
        /// Gets or sets the list of rows within the table.
        /// </summary>
        public List<Row> Rows { get; set; }

        /// <summary>
        /// Gets or sets the column properties of the table.
        /// </summary>
        public Column Column { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Table"/> class with empty rows and default column properties.
        /// </summary>
        public Table()
        {
            Rows = new List<Row>();
            Column = new Column();
            Style = "TableGrid";
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Table"/> class with a specified number of rows and columns.
        /// </summary>
        /// <param name="rows">The number of rows in the table.</param>
        /// <param name="columns">The number of columns in the table.</param>
        public Table(int rows, int columns)
        {
            Rows = new List<Row>();
            Column = new Column();

            for (var i = 0; i < rows; i++)
            {
                var row = new Row();
                row.Cells = new List<Cell>();

                for (var j = 0; j < columns; j++)
                {
                    var cellContent = "";
                    var paragraph = new Paragraph();
                    paragraph.AddRun(new Run { Text = cellContent });

                    var cell = new Cell { Paragraphs = new List<Paragraph> { paragraph } };
                    row.Cells.Add(cell);
                }

                Rows.Add(row);
            }
        }
    }
    /// <summary>
    /// Represents a row within a table in a Word document.
    /// </summary>
    public class Row
    {
        /// <summary>
        /// Gets or sets the list of cells within the row.
        /// </summary>
        public List<Cell> Cells { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Row"/> class with empty cells.
        /// </summary>
        public Row()
        {
            Cells = new List<Cell>();
        }
    }

    /// <summary>
    /// Represents a cell within a row of a table in a Word document.
    /// </summary>
    public class Cell
    {
        /// <summary>
        /// Gets or sets the list of paragraphs within the cell.
        /// </summary>
        public List<Paragraph> Paragraphs { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Cell"/> class with empty paragraphs.
        /// </summary>
        public Cell()
        {
            Paragraphs = new List<Paragraph>();
        }
    }

    /// <summary>
    /// Represents column properties of a table in a Word document.
    /// </summary>
    public class Column
    {
        /// <summary>
        /// Gets or sets the width of the column.
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Column"/> class with a default width of 0.
        /// </summary>
        public Column()
        {
            Width = 0;
        }
    }

    /// <summary>
    /// Represents a section element in a Word document.
    /// </summary>
    public class Section : IElement
    {
        /// <summary>
        /// Gets the unique identifier of the section.
        /// </summary>
        public int ElementId { get; internal set; }

        /// <summary>
        /// Gets the page size properties for the section.
        /// </summary>
        public PageSize PageSize { get; internal set; }

        /// <summary>
        /// Gets the page margin properties for the section.
        /// </summary>
        public PageMargin PageMargin { get; internal set; }
        internal Section()
        {
            //Do nothing
        }
    }

    /// <summary>
    /// Represents the page size properties of a section in a Word document.
    /// </summary>
    public class PageSize
    {
        /// <summary>
        /// Gets sets the height of the page.
        /// </summary>
        public int Height { get; internal set; }

        /// <summary>
        /// Gets the width of the page.
        /// </summary>
        public int Width { get; internal set; }

        /// <summary>
        /// Gets the orientation of the page (e.g., "Portrait" or "Landscape").
        /// </summary>
        public string Orientation { get; internal set; }
        internal PageSize()
        {
            //Do nothing
        }
    }

    /// <summary>
    /// Represents the page margin properties of a section in a Word document.
    /// </summary>
    public class PageMargin
    {
        /// <summary>
        /// Gets the top margin of the page.
        /// </summary>
        public int Top { get; internal set; }

        /// <summary>
        /// Gets the right margin of the page.
        /// </summary>
        public int Right { get; internal set; }

        /// <summary>
        /// Gets the bottom margin of the page.
        /// </summary>
        public int Bottom { get; internal set; }

        /// <summary>
        /// Gets the left margin of the page.
        /// </summary>
        public int Left { get; internal set; }

        /// <summary>
        /// Gets the header margin of the page.
        /// </summary>
        public int Header { get; internal set; }

        /// <summary>
        /// Gets the footer margin of the page.
        /// </summary>
        public int Footer { get; internal set; }
        internal PageMargin()
        {
            //Do nothing
        }
    }

    /// <summary>
    /// Represents an unknown element in a Word document.
    /// </summary>
    public class Unknown : IElement
    {
        /// <summary>
        /// Gets the unique identifier of the unknown element.
        /// </summary>
        public int ElementId { get; internal set; }

        internal Unknown()
        {
            // Do nothing
        }
    }
    /// <summary>
    /// Represents Styles associated with different elements.
    /// </summary>
    public class ElementStyles
    {
        /// <summary>
        /// Gets the fonts defined in theme.
        /// </summary>
        public List<string> ThemeFonts { get; internal set; }
        /// <summary>
        /// Gets the fonts defined in FontTable
        /// </summary>
        public List<string> TableFonts { get; internal set; }
        /// <summary>
        /// Gets the Paragraph Styles
        /// </summary>
        public List<string> ParagraphStyles { get; internal set; }
        /// <summary>
        /// Gets the Table Styles
        /// </summary>
        public List<string> TableStyles { get; internal set; }
        /// <summary>
        /// Initializes all Styles.
        /// </summary>
        public ElementStyles()
        {
            ThemeFonts = new List<string>();
            TableFonts = new List<string>();
            ParagraphStyles = new List<string>();
            TableStyles = new List<string>();
        }
    }

}