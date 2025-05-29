using Openize.Slides.Common.Enumerations;
using System;
using System.Collections.Generic;
using System.Text;

namespace Openize.Slides.Common
{   /// <summary>
    /// This class provides essential static methods for generating unique relationship IDs, obtaining random slide IDs, and converting measurements.
    /// </summary>
    public static class Utility
    {
        // Example of a static variable to keep track of the next index
        private static int nextIndex = 0;
        private static int slideNextIndex = 0;
        /// <summary>
        /// Property to set next index for slide relationship Id.
        /// </summary>
        public static int NextIndex { get => nextIndex; set => nextIndex = value; }
        public static int SlideNextIndex { get => nextIndex; set => nextIndex = value; }

        /// <summary>
        /// Function to generate a unique Relationship ID
        /// </summary>
        /// <returns></returns>
        public static string GetUniqueRelationshipId ()
        {
            // Assuming a starting index of 2, you can modify this based on your needs
            int nextIndex = NextIndex;
            return $"rId{nextIndex}";
        }

        /// <summary>
        /// Function to get unique slide Id.
        /// </summary>
        /// <returns></returns>
        public static uint GetRandomSlideId ()
        {
            // You can implement your logic to generate a random ID here.
            // For simplicity, I'll use a simple random number for illustration.
            Random random = new Random();
            return (uint)random.Next(1, int.MaxValue);
        }
        /// <summary>
        /// Function to convert EMU to Pixel
        /// </summary>
        /// <param name="emuValue">Long value</param>
        /// <returns></returns>
       public static double EmuToPixels (long emuValue)
        {
            const double emuPerInch = 914400.0;
            const double pixelsPerInch = 96.0; // Standard screen resolution

            return emuValue * pixelsPerInch / emuPerInch;
        }
        /// <summary>
        /// Function to convert Pixel valie to EMU.
        /// </summary>
        /// <param name="pixelsValue">Double value</param>
        /// <returns></returns>
       public static long PixelsToEmu (double pixelsValue)
        {
            const double emuPerInch = 914400.0;
            const double pixelsPerInch = 96.0; // Standard screen resolution

            return (long)(pixelsValue * emuPerInch / pixelsPerInch);
        }
        public static Stylings DeserializeStyling(string stylingInfo)
        {
            // Here you would deserialize the styling information from the string format
            // For example, JSON deserialization
            return Newtonsoft.Json.JsonConvert.DeserializeObject<Stylings>(stylingInfo);
        }
        public static string SerializeStyling(Stylings styling)
        {
            // Here you would serialize the styling object into a string format
            // For simplicity, let's assume it's JSON serialization
            return Newtonsoft.Json.JsonConvert.SerializeObject(styling);
        }

    }
    /// <summary>
    /// Common class to get the hexadecimal values of colors as string.
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
    /// Custom exception class for file format-related exceptions.
    /// </summary>
    public class OpenizeException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="OpenizeException"/> class with a specified error message and a reference to the inner exception.
        /// </summary>
        /// <param name="message">The error message that explains the reason for the exception.</param>
        /// <param name="innerException">The exception that is the cause of the current exception, or a null reference if no inner exception is specified.</param>
        public OpenizeException (string message, Exception innerException) : base(message, innerException)
        {
            //Do nothing
        }

        public static string ConstructMessage (Exception Ex, string Operation)
        {
            return $"Error in Operation {Operation} at Openize.OpenXML-SDK: {Ex.Message} \n Inner Exception: {Ex.InnerException?.Message ?? "N/A"}";
        }
    }

    /// <summary>
    /// Represents the stylings applied to text elements.
    /// </summary>
    public struct Stylings
    {
        private int _fontSize;
        private TextAlignment _alignment;
        private string _fontFamily;
        private string _textColor;

        /// <summary>
        /// Initializes a new instance of the Stylings struct with default values.
        /// </summary>
        /// <param name="fontSize">The font size (default is 12).</param>
        /// <param name="alignment">The text alignment (default is TextAlignment.Left).</param>
        /// <param name="fontFamily">The font family (default is "Calibri").</param>
        /// <param name="textColor">The text color in hexadecimal format (default is "#000000").</param>
        public Stylings(int fontSize = 12, TextAlignment alignment = TextAlignment.Left, string fontFamily = "Calibri", string textColor = "#000000")
        {
            _fontSize = fontSize;
            _alignment = alignment;
            _fontFamily = fontFamily;
            _textColor = textColor;
        }

        /// <summary>
        /// Gets or sets the font size.
        /// </summary>
        public int FontSize
        {
            get => _fontSize;
            set => _fontSize = value;
        }

        /// <summary>
        /// Gets or sets the text alignment.
        /// </summary>
        public TextAlignment Alignment
        {
            get => _alignment;
            set => _alignment = value;
        }

        /// <summary>
        /// Gets or sets the font family.
        /// </summary>
        public string FontFamily
        {
            get => _fontFamily;
            set => _fontFamily = value ?? "Calibri"; // Set default value if null 
        }

        /// <summary>
        /// Gets or sets the text color in hexadecimal format.
        /// </summary>
        public string TextColor
        {
            get => _textColor;
            set => _textColor = value ?? "#000000"; // Set default value if null 
        }
    }

}
