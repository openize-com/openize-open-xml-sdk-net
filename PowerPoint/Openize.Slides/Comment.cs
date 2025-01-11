using Openize.Slides.Facade;
using System;
using System.Collections.Generic;
using System.Text;

namespace Openize.Slides
{
    /// <summary>
    /// Represents a comment within the system.
    /// </summary>
    public class Comment
    {
        /// <summary>
        /// Reference to the CommentFacade object.
        /// </summary>
        private CommentFacade _Facade;

        /// <summary>
        /// Index of the comment.
        /// </summary>
        private int _CommentIndex = 0;

        /// <summary>
        /// Content of the comment.
        /// </summary>
        private String _Text;

        /// <summary>
        /// X-coordinate of the comment.
        /// </summary>
        private long _x;

        /// <summary>
        /// Y-coordinate of the comment.
        /// </summary>
        private long _y;

        /// <summary>
        /// Time when the comment was inserted.
        /// </summary>
        private DateTime _InsertedAt;

        /// <summary>
        /// ID of the comment's author.
        /// </summary>
        private int _AuthorId;

        /// <summary>
        /// Gets or sets the CommentFacade object.
        /// </summary>
        public CommentFacade Facade { get => _Facade; set => _Facade = value; }

        /// <summary>
        /// Gets or sets the index of the comment.
        /// </summary>
        public int CommentIndex { get => _CommentIndex; set => _CommentIndex = value; }

        /// <summary>
        /// Gets or sets the content of the comment.
        /// </summary>
        public string Text { get => _Text; set => _Text = value; }

        /// <summary>
        /// Gets or sets the X-coordinate of the comment.
        /// </summary>
        public long X { get => _x; set => _x = value; }

        /// <summary>
        /// Gets or sets the Y-coordinate of the comment.
        /// </summary>
        public long Y { get => _y; set => _y = value; }

        /// <summary>
        /// Gets or sets the time when the comment was inserted.
        /// </summary>
        public DateTime InsertedAt { get => _InsertedAt; set => _InsertedAt = value; }

        /// <summary>
        /// Gets or sets the ID of the comment's author.
        /// </summary>
        public int AuthorId { get => _AuthorId; set => _AuthorId = value; }

        /// <summary>
        /// Initializes a new instance of the Comment class.
        /// </summary>
        public Comment()
        {
            //_Facade = new CommentFacade();
        }

        /// <summary>
        /// Method to remove the comment.
        /// </summary>
        public void Remove()
        {
            _Facade.RemoveComment(_CommentIndex);
        }
    }

}
