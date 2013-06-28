using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Excel.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelViewProvider : IViewProvider<ExcelRibbonType>
    {
        private readonly Dictionary<Workbook, List<Window>> documents;
        private Application excelApplication;
        private readonly WorkbookClosedMonitor _monitor;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelViewProvider"/> class.
        /// </summary>
        /// <param name="excelApplication">The Excel application.</param>
        public ExcelViewProvider(Application excelApplication)
        {
            documents = new Dictionary<Workbook, List<Window>>();
            this.excelApplication = excelApplication;
            _monitor = new WorkbookClosedMonitor(excelApplication);
            _monitor.WorkbookClosed += MonitorWorkbookClosed;
        }

        void MonitorWorkbookClosed(object sender, WorkbookClosedEventArgs e)
        {
            var handler = ViewClosed;
            if (handler == null) return;

            var windows = documents[e.Workbook];

            foreach (var window in windows)
            {
                handler(this, new ViewClosedEventArgs(window, e.Workbook));
                window.ReleaseComObject();
            }
            documents.Remove(e.Workbook);
        }

        void ExcelApplicationWindowActivate(Workbook doc, Window wn)
        {
            var handler = NewView;
            if (handler == null) return;
            if (!documents.ContainsKey(doc))
            {
                documents.Add(doc, new List<Window>());
            }

            //Check if we have this window registered
            if (documents[doc].Contains(wn)) return;

            documents[doc].Add(wn);
            handler(this, new NewViewEventArgs<ExcelRibbonType>(wn, doc, ExcelRibbonType.ExcelWorkbook));
        }

        /// <summary>
        /// Initialises this instance.
        /// </summary>
        public void Initialise()
        {
            excelApplication.WindowActivate += ExcelApplicationWindowActivate;
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
            excelApplication.WindowActivate -= ExcelApplicationWindowActivate;
            excelApplication = null;
        }

        /// <summary>
        /// Registers the open Excel documents.
        /// </summary>
        public void RegisterOpenDocuments()
        {
            using (var documents = excelApplication.Workbooks.WithComCleanup())
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