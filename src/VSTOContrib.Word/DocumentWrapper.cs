using System;
using Microsoft.Office.Interop.Word;

namespace VSTOContrib.Word
{
    /// <summary>
    /// 
    /// </summary>
    public class DocumentWrapper
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentWrapper"/> class.
        /// </summary>
        /// <param name="document">The document.</param>
        public DocumentWrapper(Document document)
        {
            Document = document;
            ((DocumentEvents2_Event)Document).Close += DocumentClose;
        }

        /// <summary>
        /// Occurs when inspector is closed.
        /// </summary>
        public event EventHandler<DocumentClosedEventArgs> Closed;

        /// <summary>
        /// Gets the inspector.
        /// </summary>
        /// <value>The inspector.</value>
        public Document Document { get; private set; }

        private void DocumentClose()
        {
            ((DocumentEvents2_Event)Document).Close -= DocumentClose; 
            
            var handler = Closed;
            if (handler != null)
                Closed(this, new DocumentClosedEventArgs(Document));

            Document = null;
        }
    }
}
