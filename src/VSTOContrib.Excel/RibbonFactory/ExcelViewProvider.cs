using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using VSTOContrib.Core;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Excel.RibbonFactory
{
    public class ExcelViewProvider : IViewProvider
    {
        readonly Dictionary<Workbook, List<Window>> workbooks;
        Application excelApplication;
        Window singleWindow;

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
            excelApplication.WorkbookOpen += OnInitialise;
        }

        void OnInitialise(Workbook wb)
        {
            if (!workbooks.ContainsKey(wb))
                workbooks.Add(wb, new List<Window>());

            if (IsMdi())
            {
                if (singleWindow == null)
                    singleWindow = wb.Windows[1];
                workbooks[wb].Add(singleWindow);
                NewView(this, new NewViewEventArgs(singleWindow, wb, ExcelRibbonType.ExcelWorkbook.GetEnumDescription()));
            }
            else
            {
                foreach (Window window in wb.Windows)
                {
                    workbooks[wb].Add(window);
                    NewView(this, new NewViewEventArgs(window, wb, ExcelRibbonType.ExcelWorkbook.GetEnumDescription()));
                }
            }

            wb.WindowActivate += wn => Activate(wb, wn);

            wb.WindowDeactivate += wn =>
            {
                if (IsMdi())
                {
                    var args = new HideCustomTaskPanesForContextEventArgs(wb, false);
                    UpdateCustomTaskPanesVisibilityForContext(this, args);
                }
            };
        }

        void Activate(Workbook wb, Window wn)
        {
            if (IsMdi() && !workbooks[wb].Contains(singleWindow))
                workbooks[wb].Add(singleWindow);
            if (!IsMdi() && !workbooks[wb].Contains(wn))
            {
                var windows = workbooks[wb];
                if (windows.All(w => ((dynamic) w).Hwnd != ((dynamic) wn).Hwnd))
                {
                    windows.Add(wn);
                    NewView(this, new NewViewEventArgs(wn, wb, ExcelRibbonType.ExcelWorkbook.GetEnumDescription()));
                }
            }

            if (IsMdi())
                UpdateCustomTaskPanesVisibilityForContext(this, new HideCustomTaskPanesForContextEventArgs(wb, true));
        }

        bool IsMdi()
        {
            return new Version(excelApplication.Version).Major <= 14;
        }

        public event EventHandler<NewViewEventArgs> NewView = (sender, args) => { };
        public event EventHandler<ViewClosedEventArgs> ViewClosed = (sender, args) => { };

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