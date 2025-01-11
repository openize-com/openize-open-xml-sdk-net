using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.IO;
using System;

namespace Openize.Cells
{

    /// <summary>
    /// Represents an image, providing methods and properties to handle its path, data, and extension.
    /// </summary>
    public class Image
    {
        /// <summary>
        /// Gets the path of the image if initialized using a file path.
        /// </summary>
        public string Path { get; }

        /// <summary>
        /// Gets the stream data of the image if initialized using a stream.
        /// </summary>
        public Stream Data { get; }

        /// <summary>
        /// Gets the file extension of the image.
        /// </summary>
        public string Extension { get; }


        /// <summary>
        /// Initializes a new instance of the <see cref="Image"/> class using a file path.
        /// </summary>
        /// <param name="path">The path to the image file.</param>
        /// <exception cref="ArgumentException">
        /// Thrown when <paramref name="path"/> is null or empty or when the file does not exist.
        /// </exception>
        public Image(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new ArgumentException("Valid path required", nameof(path));

            Path = path;
            Extension = System.IO.Path.GetExtension(path);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Image"/> class using a stream and a file extension.
        /// </summary>
        /// <param name="data">The stream containing the image data.</param>
        /// <param name="extension">The file extension of the image.</param>
        /// <exception cref="ArgumentNullException">
        /// Thrown when <paramref name="data"/> or <paramref name="extension"/> is null.
        /// </exception>
        public Image(Stream data, string extension)
        {
            Data = data ?? throw new ArgumentNullException(nameof(data));
            Extension = extension ?? throw new ArgumentNullException(nameof(extension));
        }

    }


    

}

