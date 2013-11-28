using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using VSTOContrib.Core;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Word.RibbonFactory
{
    public class WordViewProvider : IViewProvider
    {
        private readonly Dictionary<Document, List<Window>> documents;
        private readonly Dictionary<Document, DocumentWrapper> documentWrappers;
        private Application wordApplication;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordViewProvider"/> class.
        /// </summary>
        /// <param name="wordApplication">The word application.</param>
        public WordViewProvider(Application wordApplication)
        {
            documentWrappers = new Dictionary<Document, DocumentWrapper>();
            documents = new Dictionary<Document, List<Window>>();
            this.wordApplication = wordApplication;
        }

        void WordApplicationWindowActivate(Document doc, Window wn)
        {
            var handler = NewView;
            if (handler == null) return;
            if (!documents.ContainsKey(doc))
            {
                documents.Add(doc, new List<Window>());
                var documentWrapper = new DocumentWrapper(doc);
                documentWrapper.Closed += DocumentClosed;
                documentWrappers.Add(doc, documentWrapper);
            }

            //Check if we have this window registered
            if (documents[doc].Contains(wn)) return;

            documents[doc].Add(wn);
            handler(this, new NewViewEventArgs(wn, doc, WordRibbonType.WordDocument.GetEnumDescription()));
        }

        void DocumentClosed(object sender, DocumentClosedEventArgs e)
        {
            var handler = ViewClosed;
            if (handler == null) return;

            documentWrappers.Remove(e.Document);
            var windows = documents[e.Document];

            foreach (var window in windows)
            {
                handler(this, new ViewClosedEventArgs(window, e.Document));
                window.ReleaseComObject();
            }
            documents.Remove(e.Document);
        }

        /// <summary>
        /// Initialises this instance.
        /// </summary>
        public void Initialise()
        {
            wordApplication.WindowActivate += WordApplicationWindowActivate;
            wordApplication.DocumentOpen += WordApplicationDocumentOpen;
            //TODO protected window activate
        }

        static void WordApplicationDocumentOpen(Document doc)
        {

        }

        /// <summary>
        /// Occurs when [new view].
        /// </summary>
        public event EventHandler<NewViewEventArgs> NewView;
        /// <summary>
        /// Occurs when [view closed].
        /// </summary>
        public event EventHandler<ViewClosedEventArgs> ViewClosed;

        /// <summary>
        /// Raise when the custom task panes for a context need to change their visibility
        /// </summary>
        public event EventHandler<HideCustomTaskPanesForContextEventArgs> UpdateCustomTaskPanesVisibilityForContext;

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
            wordApplication.WindowActivate -= WordApplicationWindowActivate;
            wordApplication = null;
        }

        /// <summary>
        /// Registers the open word documents.
        /// </summary>
        public void RegisterOpenDocuments()
        {
            using (var documents = wordApplication.Documents.WithComCleanup())
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