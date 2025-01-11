using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using Text = DocumentFormat.OpenXml.Drawing.Text;
using NonVisualGroupShapeProperties = DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using Openize.Slides.Common.Enumerations;
using Openize.Slides.Common;
using System.Collections.Generic;
using System;

namespace Openize.Slides.Facade
{
    
    public class ListFacade
    {
        private ListStyle _ListStyle;
        private ListType _ListType;
        private List<string> _ListItems;
        private String _TextColor;
        private String _FontFamily;
        private int _FontSize;
        private P.Shape _TextShape;
        private AnimationType _Animation = AnimationType.None;

        public ListType ListType { get => _ListType; set => _ListType = value; }
        public List<string> ListItems { get => _ListItems; set => _ListItems = value; }
        public string TextColor { get => _TextColor; set => _TextColor = value; }
        public string FontFamily { get => _FontFamily; set => _FontFamily = value; }
        public int FontSize { get => _FontSize; set => _FontSize = value; }
        public P.Shape TextShape { get => _TextShape; set => _TextShape = value; }
        public AnimationType Animation { get => _Animation; set => _Animation = value; }

        public ListFacade ()
        {
            _ListStyle = new ListStyle();
        }
        private Paragraph AddListItem (String Text, String _textColor, String _fontFamily, int _fontSize)
        {
            _TextColor = _textColor;
            _FontFamily = _fontFamily;
            _FontSize = _fontSize;
            Paragraph paragraph1 = new Paragraph();

            ParagraphProperties paragraphProperties1 = new ParagraphProperties() { LeftMargin = 342900, Indent = -342900, Alignment = TextAlignmentTypeValues.Left };
            BulletFont bulletFont1 = new BulletFont() { Typeface = _fontFamily, Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };

            if(_ListType == ListType.Bulleted){
                CharacterBullet characterBullet1 = new CharacterBullet() { Char = "•" };
                paragraphProperties1.Append(bulletFont1);
                paragraphProperties1.Append(characterBullet1); 
            }else if(_ListType == ListType.Numbered)
            {
                AutoNumberedBullet autoNumberedBullet1 = new AutoNumberedBullet() { Type = TextAutoNumberSchemeValues.ArabicPeriod };

                paragraphProperties1.Append(bulletFont1);
                paragraphProperties1.Append(autoNumberedBullet1);
            }

            Run run1 = new Run();
            RunProperties runProperties1 = new RunProperties(new SolidFill(new RgbColorModelHex() { Val = TextColor }),
                 new LatinFont() { Typeface = _fontFamily }) { FontSize = _fontSize, Language = "en-US", Dirty = false };
            Text text1 = new Text();
            text1.Text = Text;

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            return paragraph1;
        }
        public P.TextBody CreateList (List<String> listItems,String _textColor, String _fontFamily, int _fontSize, P.TextBody body)
        {
            _ListItems = listItems;
            body.Append(new D.BodyProperties() { RightToLeftColumns = false, Anchor = D.TextAnchoringTypeValues.Center });
            body.Append(_ListStyle);
            foreach(var text in listItems)
            {
                body.Append(AddListItem(text, _textColor, _fontFamily, _fontSize));
            }
            return body;
        }
        public void Update ()
        {
            var listItems = _ListItems;
            var body = _TextShape.TextBody;
            body.RemoveAllChildren<Paragraph>();
            foreach (var text in listItems)
            {
                body.Append(AddListItem(text, _TextColor, _FontFamily, _FontSize));
            }
           
        }

    }
}
