using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Office.Contrib.Extensions;
using Office.Contrib.RibbonFactory;
using Office.Contrib.RibbonFactory.Interfaces;

namespace Office.Excel.Contrib.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelViewProvider : IViewProvider<ExcelRibbonType>
    {
        private readonly Dictionary<Workbook, List<Window>> _documents;
        private Application _excelApplication;
        private readonly WorkbookClosedMonitor _monitor;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelViewProvider"/> class.
        /// </summary>
        /// <param name="excelApplication">The Excel application.</param>
        public ExcelViewProvider(Application excelApplication)
        {
            _documents = new Dictionary<Workbook, List<Window>>();
            _excelApplication = excelApplication;
            _monitor = new WorkbookClosedMonitor(excelApplication);
            _monitor.WorkbookClosed += MonitorWorkbookClosed;
        }

        void MonitorWorkbookClosed(object sender, WorkbookClosedEventArgs e)
        {
            var handler = ViewClosed;
            if (handler == null) return;

            var windows = _documents[e.Workbook];

            foreach (var window in windows)
            {
                handler(this, new ViewClosedEventArgs(window, e.Workbook));
                window.ReleaseComObject();
            }
            _documents.Remove(e.Workbook);
        }

        void ExcelApplicationWindowActivate(Workbook doc, Window wn)
        {
            var handler = NewView;
            if (handler == null) return;
            if (!_documents.ContainsKey(doc))
            {
                _documents.Add(doc, new List<Window>());
            }

            //Check if we have this window registered
            if (_documents[doc].Contains(wn)) return;

            _documents[doc].Add(wn);
            handler(this, new NewViewEventArgs<ExcelRibbonType>(wn, doc, ExcelRibbonType.ExcelWorkbook));
        }

        /// <summary>
        /// Initialises this instance.
        /// </summary>
        public void Initialise()
        {
            _excelApplication.WindowActivate += ExcelApplicationWindowActivate;
        }

        /// <summary>
        /// Occurs when [new view].
        /// </summary>
        public event EventHandler<NewViewEventArgs<ExcelRibbonType>> NewView;
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
            _excelApplication.WindowActivate -= ExcelApplicationWindowActivate;
            _excelApplication = null;
        }

        /// <summary>
        /// Registers the open Excel documents.
        /// </summary>
        public void RegisterOpenDocuments()
        {
            using (var documents = _excelApplication.Workbooks.WithComCleanup())
            {
                foreach (Workbook document in documents.Resource)
                {
                    using (var windows = document.Windows.WithComCleanup())
                    {
                        foreach (Window window in windows.Resource)
                        {
                            ExcelApplicationWindowActivate(document, window);
                        }
                    }
                }
            }
        }
    }
}