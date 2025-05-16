using System;
//using System;
//using System.IO;
using System.Linq;
using System.Collections.Generic;
using DF = DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using WP = DocumentFormat.OpenXml.Wordprocessing;
//using A = DocumentFormat.OpenXml.Drawing;
//using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
//using DWG = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
//using DWS = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
//using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using FF = Openize.Words.IElements;
using OWD = OpenXML.Words.Data;
//using OT = OpenXML.Templates;
//using Openize.Words;
//using System.Xml.Linq;

namespace OpenXML.Words
{
    internal class OoxmlParagraph
    {
        private readonly object _lockObject = new object();
        private List<int> _IDs;
        private NumberingDefinitionsPart _numberingPart;

        private OoxmlParagraph(List<int> IDs, NumberingDefinitionsPart numberingPart)
        {
            _IDs = IDs;
            _numberingPart = numberingPart;
        }

        internal static OoxmlParagraph CreateInstance(List<int> IDs, NumberingDefinitionsPart numberingPart)
        {
            return new OoxmlParagraph(IDs,numberingPart);
        }

        internal WP.Paragraph CreateParagraph(FF.Paragraph ffP)
        {
            lock (_lockObject)
            {
                try
                {
                    var wpParagraph = new WP.Paragraph();

                    if (ffP.Style != null)
                    {
                        var paragraphProperties = new WP.ParagraphProperties();

                        var paragraphStyleId = new WP.ParagraphStyleId { Val = ffP.Style };
                        paragraphProperties.Append(paragraphStyleId);

                        #region Create List Paragraph

                        if (ffP.Style == "ListParagraph")
                        {

                            // Check if NumberingId already exists

                            var isExist = false;
                            if (_IDs != null)
                            {
                                foreach (var id in _IDs)
                                {
                                    if (id == ffP.NumberingId)
                                    {
                                        isExist = true;

                                        var numbering = _numberingPart.Numbering;
                                        var abstractNum = numbering.Elements<WP.AbstractNum>().FirstOrDefault(an => an.AbstractNumberId == ffP.NumberingId);
                                        if (abstractNum != null)
                                        {
                                            var level = abstractNum.Elements<WP.Level>().FirstOrDefault(l => l.LevelIndex == ffP.NumberingLevel - 1);
                                            if (level != null)
                                            {
                                                if (ffP.IsAlphabeticNumber)
                                                {
                                                    level.NumberingFormat.Val = WP.NumberFormatValues.LowerLetter;
                                                    level.LevelText.Val = string.Format("%{0}.", (int)ffP.NumberingLevel);
                                                }
                                                else if (ffP.IsRoman)
                                                {
                                                    level.NumberingFormat.Val = WP.NumberFormatValues.LowerRoman;
                                                    level.LevelText.Val = string.Format("%{0}.", (int)ffP.NumberingLevel);
                                                }
                                                else if (ffP.IsBullet)
                                                {
                                                    level.NumberingFormat.Val = WP.NumberFormatValues.Bullet;
                                                    level.LevelText.Val = "o";
                                                }
                                                if (ffP.IsNumbered)
                                                {
                                                    level.NumberingFormat.Val = WP.NumberFormatValues.Decimal;
                                                    level.LevelText.Val = string.Format("%{0}.", string.Join(".%", Enumerable.Range(1, (int)ffP.NumberingLevel)));
                                                }
                                                numbering.Save();
                                            }
                                        }
                                    }
                                }
                            }
                            if (!isExist)
                            {
                                if (ffP.NumberingId != null)
                                {
                                    if (ffP.NumberingLevel == null)
                                        ffP.NumberingLevel = 1;
                                    if (ffP.IsAlphabeticNumber == false && ffP.IsBullet == false &&
                                        ffP.IsNumbered == false && ffP.IsRoman == false)
                                        ffP.IsNumbered = true;

                                    var abstractNum = new WP.AbstractNum() { AbstractNumberId = ffP.NumberingId };
                                    abstractNum.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                                    var multiLevelType = new WP.MultiLevelType() { Val = WP.MultiLevelValues.Multilevel };
                                    multiLevelType.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                                    abstractNum.Append(multiLevelType);

                                    var level = new WP.Level() { LevelIndex = 0 };

                                    for (var i = 1; i <= 9; i++)
                                    {
                                        level = new WP.Level() { LevelIndex = i - 1 };
                                        level.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                                        var numberingFormat = new WP.NumberingFormat();
                                        var levelText = new WP.LevelText();

                                        if (ffP.IsNumbered)
                                        {
                                            numberingFormat.Val = WP.NumberFormatValues.Decimal;
                                            levelText.Val = string.Format("%{0}.", string.Join(".%", Enumerable.Range(1, i)));
                                        }
                                        else if (ffP.IsAlphabeticNumber)
                                        {
                                            numberingFormat.Val = WP.NumberFormatValues.LowerLetter;
                                            levelText.Val = string.Format("%{0}.", i);
                                        }
                                        else if (ffP.IsRoman)
                                        {
                                            numberingFormat.Val = WP.NumberFormatValues.LowerRoman;
                                            levelText.Val = string.Format("%{0}.", i);
                                        }
                                        else if (ffP.IsBullet)
                                        {
                                            numberingFormat.Val = WP.NumberFormatValues.Bullet;
                                            levelText.Val = "o";
                                        }

                                        var previousParagraphProperties = new WP.PreviousParagraphProperties();
                                        var indentation = new WP.Indentation() { Left = (i * 720).ToString(), Hanging = "360" };
                                        previousParagraphProperties.Append(indentation);

                                        level.Append(new WP.StartNumberingValue() { Val = 1 });
                                        level.Append(numberingFormat);
                                        level.Append(levelText);
                                        level.Append(new WP.LevelJustification() { Val = WP.LevelJustificationValues.Left });
                                        level.Append(previousParagraphProperties);

                                        abstractNum.Append(level);
                                    }

                                    var numberingInstance = new WP.NumberingInstance() { NumberID = ffP.NumberingId };
                                    var abstractNumId = new WP.AbstractNumId() { Val = ffP.NumberingId };

                                    numberingInstance.Append(abstractNumId);
                                    _numberingPart.Numbering.Append(abstractNum);
                                    _numberingPart.Numbering.Append(numberingInstance);


                                    _IDs.Add((int)ffP.NumberingId);
                                }
                            }

                            var numberingProperties = new WP.NumberingProperties();
                            var numberingLevelReference = new WP.NumberingLevelReference() { Val = ffP.NumberingLevel - 1 };
                            var numberingId = new WP.NumberingId() { Val = ffP.NumberingId };
                            numberingProperties.Append(numberingLevelReference);
                            numberingProperties.Append(numberingId);
                            paragraphProperties.Append(numberingProperties);
                        }
                        #endregion


                        // Create Borders
                        if (ffP.ParagraphBorder.Size > 0)
                        {
                            WP.ParagraphBorders paragraphBorders = new WP.ParagraphBorders();
                            WP.TopBorder topBorder = new WP.TopBorder()
                            {
                                Val = CreateBorder(ffP.ParagraphBorder.Width),
                                Color = ffP.ParagraphBorder.Color,
                                Size = (DF.UInt32Value)(uint)ffP.ParagraphBorder.Size,
                                Space = (DF.UInt32Value)(uint)ffP.ParagraphBorder.Size
                            };
                            WP.LeftBorder leftBorder = new WP.LeftBorder()
                            {
                                Val = CreateBorder(ffP.ParagraphBorder.Width),
                                Color = ffP.ParagraphBorder.Color,
                                Size = (DF.UInt32Value)(uint)ffP.ParagraphBorder.Size,
                                Space = (DF.UInt32Value)(uint)ffP.ParagraphBorder.Size
                            };
                            WP.BottomBorder bottomBorder = new WP.BottomBorder()
                            {
                                Val = CreateBorder(ffP.ParagraphBorder.Width),
                                Color = ffP.ParagraphBorder.Color,
                                Size = (DF.UInt32Value)(uint)ffP.ParagraphBorder.Size,
                                Space = (DF.UInt32Value)(uint)ffP.ParagraphBorder.Size
                            };
                            WP.RightBorder rightBorder = new WP.RightBorder()
                            {
                                Val = CreateBorder(ffP.ParagraphBorder.Width),
                                Color = ffP.ParagraphBorder.Color,
                                Size = (DF.UInt32Value)(uint)ffP.ParagraphBorder.Size,
                                Space = (DF.UInt32Value)(uint)ffP.ParagraphBorder.Size
                            };

                            paragraphBorders.Append(topBorder);
                            paragraphBorders.Append(leftBorder);
                            paragraphBorders.Append(bottomBorder);
                            paragraphBorders.Append(rightBorder);

                            paragraphProperties.Append(paragraphBorders);
                        }
                        // Create Justification
                        WP.JustificationValues justificationValue = CreateJustification(ffP.Alignment);
                        paragraphProperties.Append(new WP.Justification { Val = justificationValue });

                        // Create Indentation
                        CreateIndentation(paragraphProperties, ffP.Indentation);

                        wpParagraph.Append(paragraphProperties);
                    }


                    foreach (var ffR in ffP.Runs)
                    {
                        var wpRun = new WP.Run();

                        var runProperties = new WP.RunProperties();

                        if (ffR.FontFamily != null)
                        {
                            var runFont = new WP.RunFonts
                            {
                                Ascii = ffR.FontFamily,
                                HighAnsi = ffR.FontFamily,
                                ComplexScript = ffR.FontFamily,
                                EastAsia = ffR.FontFamily
                            };
                            runProperties.Append(runFont);
                        }

                        if (ffR.Color != null)
                        {
                            var color = new WP.Color { Val = ffR.Color };
                            runProperties.Append(color);
                        }

                        if (ffR.FontSize > 0)
                        {
                            var fontSize = new WP.FontSize { Val = (ffR.FontSize * 2).ToString() };
                            runProperties.Append(fontSize);
                        }

                        if (ffR.Bold)
                        {

                            runProperties.Append(new WP.Bold() { Val = new DF.OnOffValue(true) });
                        }

                        if (ffR.Italic)
                        {
                            runProperties.Append(new WP.Italic());
                        }

                        if (ffR.Underline)
                        {
                            var underline = new WP.Underline { Val = WP.UnderlineValues.Single };
                            runProperties.Append(underline);
                        }

                        var text = new WP.Text(ffR.Text);
                        wpRun.Append(runProperties, text);
                        wpParagraph.AppendChild(wpRun);
                    }

                    return wpParagraph;
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Create Paragraph");
                    throw new Openize.Words.OpenizeException(errorMessage, ex);
                }
            }
        }

        private WP.JustificationValues CreateJustification(FF.ParagraphAlignment alignment)
        {
            switch (alignment)
            {
                case FF.ParagraphAlignment.Left:
                    return WP.JustificationValues.Left;
                case FF.ParagraphAlignment.Center:
                    return WP.JustificationValues.Center;
                case FF.ParagraphAlignment.Right:
                    return WP.JustificationValues.Right;
                case FF.ParagraphAlignment.Justify:
                    return WP.JustificationValues.Both;
                default:
                    return WP.JustificationValues.Left;
            }
        }

        private WP.BorderValues CreateBorder(FF.BorderWidth borderWidth)
        {
            switch (borderWidth)
            {
                case FF.BorderWidth.Single:
                    return WP.BorderValues.Single;
                case FF.BorderWidth.Double:
                    return WP.BorderValues.Double;
                case FF.BorderWidth.Dotted:
                    return WP.BorderValues.Dotted;
                case FF.BorderWidth.DotDash:
                    return WP.BorderValues.DotDash;
                default:
                    return WP.BorderValues.Single;
            }
        }

        private void CreateIndentation(WP.ParagraphProperties paragraphProperties, FF.Indentation ffIndentation)
        {
            var indentation = new WP.Indentation();

            if (ffIndentation.Left > 0)
            {
                indentation.Left = (ffIndentation.Left * 1440).ToString();
            }

            if (ffIndentation.Right > 0)
            {
                indentation.Right = (ffIndentation.Right * 1440).ToString();
            }

            if (ffIndentation.FirstLine > 0)
            {
                indentation.FirstLine = (ffIndentation.FirstLine * 1440).ToString();
            }

            if (ffIndentation.Hanging > 0)
            {
                indentation.Hanging = (ffIndentation.Hanging * 1440).ToString();
            }

            paragraphProperties.Append(indentation);
        }

        internal FF.Paragraph LoadParagraph(WP.Paragraph wpPara, int id)
        {
            lock (_lockObject)
            {
                try
                {
                    var ffP = new FF.Paragraph { ElementId = id };

                    var paraProps = wpPara.GetFirstChild<WP.ParagraphProperties>();
                    if (paraProps != null)
                    {
                        var paraStyleId = paraProps.Elements<WP.ParagraphStyleId>().FirstOrDefault();
                        if (paraStyleId != null)
                        {
                            if (paraStyleId.Val != null) ffP.Style = paraStyleId.Val.Value;
                        }
                    }

                    if (ffP.Style == "ListParagraph")
                    {
                        if (isNumbered(paraProps))
                        {
                            if (_numberingPart != null)
                            {
                                if (paraProps.NumberingProperties.NumberingId.Val != null &&
                                paraProps.NumberingProperties.NumberingLevelReference.Val != null)
                                {
                                    ffP.NumberingId = paraProps.NumberingProperties.NumberingId.Val;
                                    ffP.NumberingLevel = paraProps.NumberingProperties.NumberingLevelReference.Val + 1;

                                    var numbering = _numberingPart.Numbering;
                                    var abstractNum = numbering.Elements<WP.AbstractNum>().FirstOrDefault(an => an.AbstractNumberId == ffP.NumberingId);
                                    if (abstractNum != null)
                                    {
                                        var level = abstractNum.Elements<WP.Level>().FirstOrDefault(l => l.LevelIndex == ffP.NumberingLevel - 1);
                                        if (level != null)
                                        {
                                            if (level.NumberingFormat.Val == WP.NumberFormatValues.Decimal)
                                                ffP.IsNumbered = true;
                                            else if (level.NumberingFormat.Val == WP.NumberFormatValues.LowerLetter)
                                                ffP.IsAlphabeticNumber = true;
                                            else if (level.NumberingFormat.Val == WP.NumberFormatValues.LowerRoman)
                                                ffP.IsRoman = true;
                                            else if (level.NumberingFormat.Val == WP.NumberFormatValues.Bullet)
                                                ffP.IsBullet = true;
                                            else
                                                ffP.IsNumbered = true;
                                        }
                                    }
                                }

                            }
                        }
                    }

                    // Load Border
                    if (isBordered(paraProps))
                    {
                        var topBorder = paraProps.ParagraphBorders?.TopBorder;
                        if (topBorder != null)
                        {
                            ffP.ParagraphBorder.Width = LoadBorder(topBorder.Val);
                            ffP.ParagraphBorder.Color = topBorder.Color;
                            ffP.ParagraphBorder.Size = (int)(uint)topBorder.Size;
                        }
                    }

                    // Load Justification
                    if (isJustified(paraProps))
                    {
                        var justificationElement = paraProps.Elements<WP.Justification>().FirstOrDefault();
                        if (justificationElement != null)
                            ffP.Alignment = LoadAlignment(justificationElement.Val);
                    }
                    else ffP.Alignment = FF.ParagraphAlignment.Left;

                    // Load Indentation
                    if (isIndented(paraProps))
                    {
                        var Indentation = paraProps.Elements<WP.Indentation>().FirstOrDefault();
                        if (Indentation != null)
                        {
                            if (Indentation.Left != null)
                                ffP.Indentation.Left = int.Parse(Indentation.Left);
                            if (Indentation.Right != null)
                                ffP.Indentation.Right = int.Parse(Indentation.Right);
                            if (Indentation.Hanging != null)
                                ffP.Indentation.Hanging = int.Parse(Indentation.Hanging);
                            if (Indentation.FirstLine != null)
                                ffP.Indentation.FirstLine = int.Parse(Indentation.FirstLine);
                        }
                    }

                    var runs = wpPara.Elements<WP.Run>();

                    foreach (var wpR in runs)
                    {
                        var fontSize = wpR.RunProperties?.FontSize?.Val != null
                            ? int.Parse(wpR.RunProperties.FontSize.Val)
                            : (int?)null;
                        if (fontSize != null) fontSize /= 2;
                        var ffR = new FF.Run
                        {
                            Text = wpR.InnerText,
                            FontFamily = wpR.RunProperties?.RunFonts?.Ascii ?? null,
                            FontSize = fontSize ?? 0,
                            Color = wpR.RunProperties?.Color?.Val ?? null,
                            Bold = (wpR.RunProperties != null && wpR.RunProperties.Bold != null),
                            Italic = (wpR.RunProperties != null && wpR.RunProperties.Italic != null),
                            Underline = (wpR.RunProperties != null && wpR.RunProperties.Underline != null)
                        };
                        ffP.AddRun(ffR);
                    }

                    return ffP;
                }

                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Load Paragraph");
                    throw new Openize.Words.OpenizeException(errorMessage, ex);
                }
            }
        }

        private FF.ParagraphAlignment LoadAlignment(WP.JustificationValues justificationValue)
        {
            if (justificationValue == WP.JustificationValues.Left)
                return FF.ParagraphAlignment.Left;
            else if (justificationValue == WP.JustificationValues.Center)
                return FF.ParagraphAlignment.Center;
            else if (justificationValue == WP.JustificationValues.Right)
                return FF.ParagraphAlignment.Right;
            else if (justificationValue == WP.JustificationValues.Both)
                return FF.ParagraphAlignment.Justify;
            else
                return FF.ParagraphAlignment.Left;
        }

        private FF.BorderWidth LoadBorder(WP.BorderValues borderValue)
        {
            if (borderValue == WP.BorderValues.Single)
                return FF.BorderWidth.Single;
            else if (borderValue == WP.BorderValues.Double)
                return FF.BorderWidth.Double;
            else if (borderValue == WP.BorderValues.Dotted)
                return FF.BorderWidth.Dotted;
            else if (borderValue == WP.BorderValues.DotDash)
                return FF.BorderWidth.DotDash;
            else
                return FF.BorderWidth.Single;
        }

        private bool isBordered(WP.ParagraphProperties prop)
        {
            try
            {
                var paragraphBorders = prop.ParagraphBorders;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private bool isJustified(WP.ParagraphProperties prop)
        {
            try
            {
                var justification = prop.Justification;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private bool isIndented(WP.ParagraphProperties prop)
        {
            try
            {
                var indentation = prop.Indentation;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private bool isNumbered(WP.ParagraphProperties prop)
        {
            try
            {
                var numbering = prop.NumberingProperties;
                var numberingId = numbering.NumberingId.Val;
                var numberingRef = numbering.NumberingLevelReference.Val;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

    }

}
