using System;
using System.Runtime.Serialization;

namespace Outlook.Utility.RibbonFactory
{
    /// <summary>
    /// Thrown when a view cannot be found for a IViewModel
    /// </summary>
    [Serializable]
    public class ViewNotFoundException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ViewNotFoundException"/> class.
        /// </summary>
        public ViewNotFoundException()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ViewNotFoundException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        public ViewNotFoundException(string message) : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ViewNotFoundException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="inner">The inner.</param>
        public ViewNotFoundException(string message, Exception inner) : base(message, inner)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ViewNotFoundException"/> class.
        /// </summary>
        /// <param name="info">The <see cref="T:System.Runtime.Serialization.SerializationInfo"/> that holds the serialized object data about the exception being thrown.</param>
        /// <param name="context">The <see cref="T:System.Runtime.Serialization.StreamingContext"/> that contains contextual information about the source or destination.</param>
        /// <exception cref="T:System.ArgumentNullException">
        /// The <paramref name="info"/> parameter is null.
        /// </exception>
        /// <exception cref="T:System.Runtime.Serialization.SerializationException">
        /// The class name is null or <see cref="P:System.Exception.HResult"/> is zero (0).
        /// </exception>
        protected ViewNotFoundException(
            SerializationInfo info,
            StreamingContext context) : base(info, context)
        {
        }
    }
}