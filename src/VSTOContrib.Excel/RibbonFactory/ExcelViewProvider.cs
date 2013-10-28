using System;
using System.Collections.Generic;
using System.Linq;
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
        readonly Dictionary<Workbook, List<Window>> workbooks;
        Application excelApplication;
        Window singleWindow;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelViewProvider"/> class.
        /// </summary>
        /// <param name="excelApplication">The Excel application.</param>
        public ExcelViewProvider(Application excelApplication)
        {
            workbooks = new Dictionary<Workbook, List<Window>>();
            this.excelApplication = excelApplication;
            var monitor = new WorkbookClosedMonitor(excelApplication);
            monitor.WorkbookClosed += MonitorWorkbookClosed;
        }

        void MonitorWorkbookClosed(object sender, WorkbookClosedEventArgs e)
        {
            var handler = ViewClosed;
            if (handler == null) return;

            if (workbooks.ContainsKey(e.Workbook))
            {
                var windows = workbooks[e.Workbook];

                foreach (var window in windows)
                {
                    handler(this, new ViewClosedEventArgs(window, e.Workbook));
                    if (!IsMdi())
                        window.ReleaseComObject();
                }
                workbooks.Remove(e.Workbook);
            }
        }

        /// <summary>
        /// Initialises this instance.
        /// </summary>
        public void Initialise()
        {
            ((AppEvents_Event)excelApplication).NewWorkbook += OnInitialise;
            ((AppEvents_Event)excelApplication).WorkbookOpen += OnInitialise;
        }

        void OnInitialise(Workbook wb)
        {
            var handler = NewView;
            if (handler == null) return;
            if (!workbooks.ContainsKey(wb))
                workbooks.Add(wb, new List<Window>());

            if (IsMdi())
            {
                if (singleWindow == null)
                    singleWindow = wb.Windows[1];
                workbooks[wb].Add(singleWindow);
                handler(this, new NewViewEventArgs<ExcelRibbonType>(singleWindow, wb, ExcelRibbonType.ExcelWorkbook));
            }
            else
            {
                foreach (Window window in wb.Windows)
                {
                    workbooks[wb].Add(window);
                    handler(this, new NewViewEventArgs<ExcelRibbonType>(window, wb, ExcelRibbonType.ExcelWorkbook));
                }
            }

            wb.WindowActivate += wn =>
            {
                if (IsMdi() && !workbooks[wb].Contains(singleWindow))
                    workbooks[wb].Add(singleWindow);
                if (!IsMdi() && !workbooks[wb].Contains(wn))
                {
                    var windows = workbooks[wb];
                    if (windows.All(w => ((dynamic)w).Hwnd != ((dynamic)wn).Hwnd))
                    {
                        windows.Add(wn);
                        handler(this, new NewViewEventArgs<ExcelRibbonType>(wn, wb, ExcelRibbonType.ExcelWorkbook));
                    }
                }

                if (IsMdi())
                    UpdateCustomTaskPanesVsibilityForContext(this, new HideCustomTaskPanesForContextEventArgs<ExcelRibbonType>(wb, true));
            };
            wb.WindowDeactivate += wn =>
            {
                if (IsMdi())
                {
                    var args = new HideCustomTaskPanesForContextEventArgs<ExcelRibbonType>(wb, false);
                    UpdateCustomTaskPanesVsibilityForContext(this, args);
                }
            };
        }

        bool IsMdi()
        {
            return new Version(excelApplication.Version).Major <= 14;
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
        /// Raise when the custom task panes for a context need to change their visibility
        /// </summary>
        public event EventHandler<HideCustomTaskPanesForContextEventArgs<ExcelRibbonType>> UpdateCustomTaskPanesVsibilityForContext;

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
            excelApplication = null;
        }

        /// <summary>
        /// Registers the open Excel workbooks.
        /// </summary>
        public void RegisterOpenDocuments()
        {
            foreach (var wb in excelApplication.Workbooks.ComLinq<Workbook>())
            {
                OnInitialise(wb);
            }
        }
    }
}