using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Word.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    public class WordViewProvider : IViewProvider<WordRibbonType>
    {
        private readonly Dictionary<Document, List<Window>> _documents;
        private readonly Dictionary<Document, DocumentWrapper> _documentWrappers;
        private Application _wordApplication;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordViewProvider"/> class.
        /// </summary>
        /// <param name="wordApplication">The word application.</param>
        public WordViewProvider(Application wordApplication)
        {
            _documentWrappers = new Dictionary<Document, DocumentWrapper>();
            _documents = new Dictionary<Document, List<Window>>();
            _wordApplication = wordApplication;
        }

        void WordApplicationWindowActivate(Document doc, Window wn)
        {
            var handler = NewView;
            if (handler == null) return;
            if (!_documents.ContainsKey(doc))
            {
                _documents.Add(doc, new List<Window>());
                var documentWrapper = new DocumentWrapper(doc);
                documentWrapper.Closed += DocumentClosed;
                _documentWrappers.Add(doc, documentWrapper);
            }

            //Check if we have this window registered
            if (_documents[doc].Contains(wn)) return;

            _documents[doc].Add(wn);
            handler(this, new NewViewEventArgs<WordRibbonType>(wn, doc, WordRibbonType.WordDocument));
        }

        void DocumentClosed(object sender, DocumentClosedEventArgs e)
        {
            var handler = ViewClosed;
            if (handler == null) return;

            _documentWrappers.Remove(e.Document);
            var windows = _documents[e.Document];

            foreach (var window in windows)
            {
                handler(this, new ViewClosedEventArgs(window, e.Document));
                window.ReleaseComObject();
            }
            _documents.Remove(e.Document);
        }

        /// <summary>
        /// Initialises this instance.
        /// </summary>
        public void Initialise()
        {
            _wordApplication.WindowActivate += WordApplicationWindowActivate;
            _wordApplication.DocumentOpen += WordApplicationDocumentOpen;
            //TODO protected window activate
        }

        static void WordApplicationDocumentOpen(Document doc)
        {

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
        /// <param name="context"></param>
        public void CleanupReferencesTo(object view, object context)
        {
            
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            _wordApplication.WindowActivate -= WordApplicationWindowActivate;
            _wordApplication = null;
        }

        /// <summary>
        /// Registers the open word documents.
        /// </summary>
        public void RegisterOpenDocuments()
        {
            using (var documents = _wordApplication.Documents.WithComCleanup())
            {
                foreach (Document document in documents.Resource)
                {
                    using (var windows = document.Windows.WithComCleanup())
                    {
                        foreach (Window window in windows.Resource)
                        {
                            WordApplicationWindowActivate(document, window);
                        }
                    }
                }
            }
        }
    }
}