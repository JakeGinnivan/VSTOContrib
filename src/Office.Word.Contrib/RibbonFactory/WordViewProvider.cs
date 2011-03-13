using System;
using System.Linq;
using Microsoft.Office.Interop.Word;
using Office.Contrib.Extensions;
using Office.Contrib.RibbonFactory;

namespace Office.Word.Contrib.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    public class WordViewProvider : IViewProvider<WordRibbonType>
    {
        private Application _wordApplication;
        private Documents _documents;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordViewProvider"/> class.
        /// </summary>
        /// <param name="wordApplication">The word application.</param>
        public WordViewProvider(Application wordApplication)
        {
            _wordApplication = wordApplication;
            wordApplication.DocumentOpen += WordApplicationDocumentOpen;

            _documents = _wordApplication.Documents;
            foreach (Document document in _documents)
            {
                WordApplicationDocumentOpen(document);
            }
        }

        void WordApplicationDocumentOpen(Document doc)
        {
            var handler = NewView;
            if (handler == null) return;

            ((DocumentEvents2_Event)doc).Close += WordViewProviderClose;

            var newViewEventArgs = new NewViewEventArgs<WordRibbonType>(doc, WordRibbonType.WordDocument);
            handler(this, newViewEventArgs);

            if (!newViewEventArgs.Handled)
                doc.ReleaseComObject();
        }

        void WordViewProviderClose()
        {
            var handler = ViewClosed;
            if (handler == null) return;

            handler(this, new ViewClosedEventArgs(_documents.Cast<object>()));
        }

        /// <summary>
        /// Occurs when [new view].
        /// </summary>
        public event EventHandler<NewViewEventArgs<WordRibbonType>> NewView;
        /// <summary>
        /// Occurs when [view closed].
        /// </summary>
        public event EventHandler<ViewClosedEventArgs> ViewClosed;

        /// <summary>
        /// Cleanups the references to a view.
        /// </summary>
        /// <param name="view">The view.</param>
        public void CleanupReferencesTo(object view)
        {
            ((DocumentEvents2_Event)view).Close -= WordViewProviderClose;
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            _documents.ReleaseComObject();
            _documents = null;
            _wordApplication.DocumentOpen -= WordApplicationDocumentOpen;
            _wordApplication = null;
        }
    }
}