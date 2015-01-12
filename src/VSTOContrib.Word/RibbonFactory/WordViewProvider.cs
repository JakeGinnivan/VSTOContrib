using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Word;
using VSTOContrib.Core;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;

namespace VSTOContrib.Word.RibbonFactory
{
    public class WordViewProvider : IViewProvider
    {
        private readonly List<int> closedDocuments = new List<int>();
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
            VstoContribLog.Info(_ => _("Application.WindowActivate raised, Document: {0}, Window: {1}", 
                doc.ToLogFormat(), wn.ToLogFormat()));
            if (!documents.ContainsKey(doc))
            {
                documents.Add(doc, new List<Window>());
                var documentWrapper = new DocumentWrapper(doc);
                documentWrapper.Closed += DocumentClosed;
                documentWrappers.Add(doc, documentWrapper);
            }

            //Check if we have this window registered
            if (documents[doc].Any(window => window == wn)) return;

            documents[doc].Add(wn);
            NewView(this, new NewViewEventArgs(wn, doc, WordRibbonType.WordDocument.GetEnumDescription()));
        }

        void DocumentClosed(object sender, DocumentClosedEventArgs e)
        {
            var document = e.Document;
            CleanupDocument(document);
        }

        void CleanupDocument(Document document)
        {
            if (!documentWrappers.ContainsKey(document)) return;

            closedDocuments.Add(document.GetHashCode());
            var documentWrapper = documentWrappers[document];
            documentWrapper.Closed -= DocumentClosed;
            documentWrappers.Remove(document);
            var windows = documents[document];

            foreach (var window in windows)
            {
                ViewClosed(this, new ViewClosedEventArgs(window, document));
                window.ReleaseComObject();
            }
            documents.Remove(document);
            if (wordApplication.Documents.Count == 1)
            {
                foreach (var viewInstance in wordApplication.Windows)
                {
                    NewView(this, new NewViewEventArgs(viewInstance, null, WordRibbonType.WordDocument.GetEnumDescription()));
                }
            }
        }

        public void Initialise()
        {
            wordApplication.WindowActivate += WordApplicationWindowActivate;
            wordApplication.DocumentOpen += WordApplicationDocumentOpen;
            wordApplication.DocumentChange += WordApplicationOnDocumentChange;
            //TODO protected window activate
        }

        void WordApplicationOnDocumentChange()
        {
            var enumDescription = WordRibbonType.WordDocument.GetEnumDescription();
            if (wordApplication.Documents.Count == 0)
            {
                VstoContribLog.Debug(_ => _("Application.DocumentChange raised, no documents currently open"));
                foreach (var viewInstance in wordApplication.Windows)
                {
                    NewView(this, new NewViewEventArgs(viewInstance, null, enumDescription));
                }
            }
            else
            {
                var activeDocument = wordApplication.ActiveDocument;
                if (closedDocuments.Contains(activeDocument.GetHashCode()))
                {
                    VstoContribLog.Debug(_ => _("Application.DocumentChange raised ActiveDocument: {0} is closing, ignoring event", activeDocument.ToLogFormat()));
                    return;
                }
                var activeWindow = wordApplication.ActiveWindow;
                VstoContribLog.Debug(_ => _("Application.DocumentChange raised, ActiveDocument: {0}, ActiveWindow: {1}",
                    activeDocument.ToLogFormat(), activeWindow.ToLogFormat()));
                NewView(this, new NewViewEventArgs(activeWindow, activeDocument, enumDescription));
            }
        }

        void WordApplicationDocumentOpen(Document doc)
        {
            VstoContribLog.Debug(_ => _("Application.DocumentOpen raised, Document: {0}", doc.ToLogFormat()));
            WordApplicationWindowActivate(doc, doc.ActiveWindow);
        }

        public event EventHandler<NewViewEventArgs> NewView = (sender, args) => { };
        public event EventHandler<ViewClosedEventArgs> ViewClosed = (sender, args) => { };
        public event EventHandler<HideCustomTaskPanesForContextEventArgs> UpdateCustomTaskPanesVisibilityForContext;

        /// <summary>
        /// Cleanups the references to a view.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <param name="context"></param>
        public void CleanupReferencesTo(object view, object context)
        {
            CleanupDocument((Document)context);
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            wordApplication.WindowActivate -= WordApplicationWindowActivate;
            wordApplication.DocumentOpen -= WordApplicationDocumentOpen;
            wordApplication.DocumentChange -= WordApplicationOnDocumentChange;
            wordApplication = null;
        }

        /// <summary>
        /// Registers the open word documents.
        /// </summary>
        public void RegisterOpenDocuments()
        {
            VstoContribLog.Debug(_ => _("Registering all already open documents"));
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