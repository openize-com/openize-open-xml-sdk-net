using Openize.Slides.Common;
using Openize.Slides.Common.Enumerations;
using Openize.Slides.Facade;
using System;
using System.Collections.Generic;

namespace Openize.Slides
{
    /// <summary>
    /// This class represents the text shape within a slide.
    /// </summary>
    public class TextShape
    {
        private String _Text;
        private Int32 _FontSize;
        private TextAlignment _Alignment = TextAlignment.None;
        private double _x;
        private double _y;
        private double _Width;
        private double _Height;
        private TextShapeFacade _Facade;
        private int _shapeIndex;
        private String _FontFamily;
        private String _TextColor;
        private List<TextSegment> _TextSegments;
        private String _BackgroundColor = null;
        private StyledList _TextList = null;
        /// <summary>
        /// Property to set or get the text of the shape.
        /// </summary>
        public string Text { get => _Text; set => _Text = value; }
        /// <summary>
        /// Property to set or get the font size of the Text Shape.
        /// </summary>
        public int FontSize { get => _FontSize; set => _FontSize = value; }
        /// <summary>
        /// Property to get or set alignment of the shape.
        /// </summary>
        public TextAlignment Alignment { get => _Alignment; set => _Alignment = value; }
        /// <summary>
        /// Property to get or set X coordinate of the shape
        /// </summary>
        public double X { get => _x; set => _x = value; }
        /// <summary>
        /// Property to get or set Y coordinate of the shape.
        /// </summary>
        public double Y { get => _y; set => _y = value; }
        /// <summary>
        /// Property to get or set width of the shape.
        /// </summary>
        public double Width { get => _Width; set => _Width = value; }
        /// <summary>
        /// Property to get or set height of the shape.
        /// </summary>
        public double Height { get => _Height; set => _Height = value; }
        /// <summary>
        /// Property to get or set the TextShapeFacade.
        /// </summary>
        public TextShapeFacade Facade { get => _Facade; set => _Facade = value; }
        /// <summary>
        /// Property to get or set the shape index within a slide.
        /// </summary>
        public int ShapeIndex { get => _shapeIndex; set => _shapeIndex = value; }
        /// <summary>
        /// Property to get or set the font family of the text shape.
        /// </summary>
        public string FontFamily { get => _FontFamily; set => _FontFamily = value; }
        /// <summary>
        /// Property to get or set the text color of the text shape.
        /// </summary>
        public string TextColor { get => _TextColor; set => _TextColor = value; }
        /// <summary>
        /// Property to set or get text segments within a text shape.
        /// </summary>
        public List<TextSegment> TextSegments { get => _TextSegments; set => _TextSegments = value; }
        /// <summary>
        /// Property to set or get background color of a text shape.
        /// </summary>
        public string BackgroundColor { get => _BackgroundColor; set => _BackgroundColor = value; }
        /// <summary>
        /// Property to set or get styled list of a text shape.
        /// </summary>
        public StyledList TextList { get => _TextList; set => _TextList = value; }



        /// <summary>
        /// Constructor of the TextShape class inititalizes the object of TextShapeFacade and populate its fields.
        /// </summary>
        public TextShape ()
        {
            _Facade = new TextShapeFacade();
            _Facade.ShapeIndex = _shapeIndex;
            _Text = "Default Text";
            _FontSize = 32;
            _FontFamily = "Calibri";
            _TextColor = "000000";
            _BackgroundColor = "Transparent";
            _Alignment = TextAlignment.Center;
            _x = Utility.EmuToPixels(1349828);
            _y = Utility.EmuToPixels(1999619);
            _Width = Utility.EmuToPixels(6000000);
            _Height = Utility.EmuToPixels(2000000);
            _TextList = null;
            Populate_Facade();
        }
        /// <summary>
        /// Method to update text shape.
        /// </summary>
        public void Update ()
        {
            Populate_Facade();
            _Facade.UpdateShape();

        }
        /// <summary>
        /// Method to populate the fields of respective facade.
        /// </summary>
        private void Populate_Facade ()
        {
            _Facade.Text = _Text;
            _Facade.FontSize = _FontSize;
            _Facade.Alignment = _Alignment;
            _Facade.TextColor = _TextColor;
            _Facade.BackgroundColor = _BackgroundColor;
            _Facade.FontFamily = _FontFamily;
            _Facade.X = Utility.PixelsToEmu(_x);
            _Facade.Y = Utility.PixelsToEmu(_y);
            _Facade.Width = Utility.PixelsToEmu(_Width);
            _Facade.Height = Utility.PixelsToEmu(_Height);
            _Facade.TextList = _TextList?.Facade;
        }
        /// <summary>
        /// Method for getting the list of text shapes.
        /// </summary>
        /// <param name="textShapeFacades">An object of TextShapeFacade.</param>
        /// <returns></returns>
        public static List<TextShape> GetTextShapes (List<TextShapeFacade> textShapeFacades)
        {
            List<TextShape> textShapes = new List<TextShape>();
            try
            {
                foreach (var facade in textShapeFacades)
                {
                    TextShape textShape = new TextShape
                    {
                        Text = facade.Text,
                        FontSize = facade.FontSize,
                        Alignment = facade.Alignment,
                        FontFamily = facade.FontFamily,
                        TextColor = facade.TextColor,
                        BackgroundColor = facade.BackgroundColor,
                        X = Utility.EmuToPixels(facade.X),
                        Y = Utility.EmuToPixels(facade.Y),
                        Width = Utility.EmuToPixels(facade.Width),
                        Height = Utility.EmuToPixels(facade.Height),
                        Facade = facade,
                        TextList = GetTextList(facade?.TextList),
                        ShapeIndex = facade.ShapeIndex
                    };

                    textShapes.Add(textShape);

                }
            }
            catch (Exception ex)
            {
                string errorMessage = Common.OpenizeException.ConstructMessage(ex, "Getting Text Shapes");
                throw new Common.OpenizeException(errorMessage, ex);
            }

            return textShapes;
        }
        /// <summary>
        /// This method is used to get the StyledList from a text shape
        /// </summary>
        /// <param name="facade"></param>
        /// <returns></returns>
        private static StyledList GetTextList (ListFacade facade)
        {
            if (facade != null)
            {
                StyledList list = new StyledList(facade.ListType);
                list.ListItems = facade.ListItems;
                list.TextColor = facade.TextColor;
                list.FontFamily = facade.FontFamily;
                list.FontSize = facade.FontSize;
                list.Facade = facade;
                return list;
            }
            return null;
        }
        /// <summary>
        /// Method to remove the textshape of a slide.
        /// </summary>
        public void Remove ()
        {
            _Facade.RemoveShape(this.Facade.TextBoxShape);
        }
    }
}
