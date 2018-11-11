// ***********************************************************************
// Assembly         : Data.FileIO.Common
// Author           : tdcart
// Created          : 04-01-2016
//
// Last Modified By : tdcart
// Last Modified On : 04-01-2016
// ***********************************************************************
// <copyright file="DuplicateColumnException.cs" company="Tim Cartwright">
//     Copyright © Tim Cartwright 2016
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Runtime.Serialization;

namespace Data.FileIO.Common
{
    /// <summary>
    /// Class DuplicateColumnException.
    /// </summary>
    /// <seealso cref="System.Exception" />
    [Serializable]
    public class DuplicateColumnNameException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DuplicateColumnNameException"/> class.
        /// </summary>
        public DuplicateColumnNameException()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DuplicateColumnNameException"/> class.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public DuplicateColumnNameException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DuplicateColumnNameException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="inner">The inner.</param>
        public DuplicateColumnNameException(string message, Exception inner)
            : base(message, inner)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DuplicateColumnNameException"/> class.
        /// </summary>
        /// <param name="info">The <see cref="T:System.Runtime.Serialization.SerializationInfo" /> that holds the serialized object data about the exception being thrown.</param>
        /// <param name="context">The <see cref="T:System.Runtime.Serialization.StreamingContext" /> that contains contextual information about the source or destination.</param>
        protected DuplicateColumnNameException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }
}
