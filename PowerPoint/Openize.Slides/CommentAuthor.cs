using Openize.Slides.Facade;
using System;
using System.Collections.Generic;
using System.Text;

namespace Openize.Slides
{
    /// <summary>
    /// Represents an author of a comment.
    /// </summary>
    public class CommentAuthor
    {
        /// <summary>
        /// Name of the author.
        /// </summary>
        private String _Name;

        /// <summary>
        /// ID of the author.
        /// </summary>
        private int _Id;

        /// <summary>
        /// Initial letter of the author's name.
        /// </summary>
        private String _InitialLetter;

        /// <summary>
        /// Color index associated with the author.
        /// </summary>
        private int _ColorIndex;

        /// <summary>
        /// Gets or sets the name of the author.
        /// </summary>
        public string Name { get => _Name; set => _Name = value; }

        /// <summary>
        /// Gets or sets the ID of the author.
        /// </summary>
        public int Id { get => _Id; set => _Id = value; }

        /// <summary>
        /// Gets or sets the initial letter of the author's name.
        /// </summary>
        public string InitialLetter { get => _InitialLetter; set => _InitialLetter = value; }

        /// <summary>
        /// Gets or sets the color index associated with the author.
        /// </summary>
        public int ColorIndex { get => _ColorIndex; set => _ColorIndex = value; }
    }

}
