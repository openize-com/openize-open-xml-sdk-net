using Openize.Slides.Common.Enumerations;
using Openize.Slides.Facade;
using System;
using System.Collections.Generic;

namespace Openize.Slides
{
    /// <summary>
    ///  This class represents the text list with bullet style.
    /// </summary>
    public class StyledList
    {
        private ListFacade _Facade;
        private ListType _ListType;
        private String _TextColor;
        private String _FontFamily;
        private int _FontSize;
        private TextShape _TextShape;
        public ListType ListType { get => _ListType; set => _ListType = value; }

        private List<String> _ListItems;
        public List<string> ListItems { get => _ListItems; set => _ListItems = value; }
        /// <summary>
        /// Property to get the facade of a styled list
        /// </summary>
        public ListFacade Facade { get => _Facade; set => _Facade = value; }
        public string TextColor { get => _TextColor; set => _TextColor = value; }
        public string FontFamily { get => _FontFamily; set => _FontFamily = value; }
        public int FontSize { get => _FontSize; set => _FontSize = value; }
        public TextShape TextShape { get => _TextShape; set => _TextShape = value; }

        /// <summary>
        /// Constructor of StyledList class.
        /// </summary>
        public StyledList (ListType type)
        {
            _Facade = new ListFacade();
            _Facade.ListType = type;
            _ListItems = new List<String>();
        }
        /// <summary>
        /// Method to add list items in styled list.
        /// </summary>
        /// <param name="text">It accepts text as list item</param>
        public void AddListItem (String text)
        {
            _ListItems.Add(text);
        }
        /// <summary>
        /// Method to update the styled list
        /// </summary>
        public void Update ()
        {
            _Facade.ListType = _ListType;
            _Facade.TextColor = _TextColor;
            _Facade.FontFamily = _FontFamily;
            _Facade.FontSize = _FontSize;
            _Facade.ListItems = _ListItems;
            _Facade.Update();
        }


    }
}
