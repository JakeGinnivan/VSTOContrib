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

            var windows = workbooks[e.Workbook];

            foreach (var window in windows)
            {
                handler(this, new ViewClosedEventArgs(window, e.Workbook));
                if (!excelApplication.ShowWindowsInTaskbar)
                    window.ReleaseComObject();
            }
            workbooks.Remove(e.Workbook);
        }

        /// <summary>
        /// Initialises this instance.
        /// </summary>
        public void Initialise()
        {
            ((AppEvents_Event)excelApplication).NewWorkbook += OnNewWorkbook;
        }

        void OnNewWorkbook(Workbook wb)
        {
            var handler = NewView;
            if (handler == null) return;
            if (!workbooks.ContainsKey(wb))
                workbooks.Add(wb, new List<Window>());

            if (excelApplication.ShowWindowsInTaskbar)
            {
                if (singleWindow == null)
                {
                    singleWindow = wb.Windows[1];
                }
                workbooks[wb].Add(singleWindow);
                handler(this, new NewViewEventArgs<ExcelRibbonType>(singleWindow, wb, ExcelRibbonType.ExcelWorkbook));
            }
            else
            {
                foreach (var window in wb.Windows.ComLinq<Window>())
                {
                    workbooks[wb].Add(window);
                    handler(this, new NewViewEventArgs<ExcelRibbonType>(window, wb, ExcelRibbonType.ExcelWorkbook));
                }
            }

            wb.WindowActivate += wn =>
            {
                if (excelApplication.ShowWindowsInTaskbar && !workbooks[wb].Contains(singleWindow))
                    workbooks[wb].Add(singleWindow);
                if (!excelApplication.ShowWindowsInTaskbar && !workbooks[wb].Contains(wn))
                    workbooks[wb].Add(wn);
                
                UpdateCustomTaskPanesVsibilityForContext(this, new HideCustomTaskPanesForContextEventArgs<ExcelRibbonType>(wb, true));
            };
            wb.WindowDeactivate += wn => UpdateCustomTaskPanesVsibilityForContext(this, new HideCustomTaskPanesForContextEventArgs<ExcelRibbonType>(wb, false));
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
                OnNewWorkbook(wb);
            }
        }
    }
}