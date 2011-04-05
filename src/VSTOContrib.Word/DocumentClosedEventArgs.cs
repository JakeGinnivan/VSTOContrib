using System;
using Microsoft.Office.Interop.Word;

namespace VSTOContrib.Word
{
    /// <summary>
    /// 
    /// </summary>
    public class DocumentClosedEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentClosedEventArgs"/> class.
        /// </summary>
        /// <param name="document">The document.</param>
        public DocumentClosedEventArgs(Document document)
        {
            Document = document;
        }

        /// <summary>
        /// Gets or sets the document.
        /// </summary>
        /// <value>The document.</value>
        public Document Document { get; set; }
    }
}