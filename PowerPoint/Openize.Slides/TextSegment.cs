using Openize.Slides.Facade;

namespace Openize.Slides
{

    /// <summary>
    /// This class represents the text segment within a paragraph.
    /// </summary>
    public class TextSegment
    {
        private int _FontSize;
        private bool _Bold;
        private bool _Italic;
        private string _FontFamily;
        private string _Color;
        private string _Text;
        private TextSegmentFacade _Facade;
        /// <summary>
        /// Property to set or get the font size of the text segment
        /// </summary>
        public int FontSize { get => _FontSize; set => _FontSize = value; }
        /// <summary>
        /// Property to make bold the text segment.
        /// </summary>
        public bool Bold { get => _Bold; set => _Bold = value; }
        /// <summary>
        /// Property to make Italic the text segment.
        /// </summary>
        public bool Italic { get => _Italic; set => _Italic = value; }
        /// <summary>
        /// Property to set font family.
        /// </summary>
        public string FontFamily { get => _FontFamily; set => _FontFamily = value; }
        /// <summary>
        /// Property to set color the text segment.
        /// </summary>
        public string Color { get => _Color; set => _Color = value; }
        /// <summary>
        /// Property to set the text of the text segment.
        /// </summary>
        public string Text { get => _Text; set => _Text = value; }
        /// <summary>
        /// Property to get facade of text segment.
        /// </summary>
        public TextSegmentFacade Facade { get => _Facade; }
        /// <summary>
        /// Method to create text segment.
        /// </summary>
        /// <returns></returns>
        public TextSegment create ()
        {
            _Facade = new TextSegmentFacade();
            Populate_Facade();
            _Facade.createTextSegment();
            return this;
        }
        /// <summary>
        /// Property to populate facade fields.
        /// </summary>
        private void Populate_Facade ()
        {
            _Facade.FontSize = _FontSize;
            _Facade.Bold = _Bold;
            _Facade.Italic = _Italic;
            _Facade.FontFamily = _FontFamily;
            _Facade.Color = _Color;
            _Facade.Text = _Text;

        }
    }

}
