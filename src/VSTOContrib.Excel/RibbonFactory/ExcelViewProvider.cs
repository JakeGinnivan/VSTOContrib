using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using VSTOContrib.Core;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;

namespace VSTOContrib.Excel.RibbonFactory
{
    public class ExcelViewProvider : IViewProvider
    {
        readonly Dictionary<Workbook, List<Window>> workbooks;
        Application excelApplication;
        Window singleWindow;
        bool nullContextOpen;

        public ExcelViewProvider(Application excelApplication)
        {
            workbooks = new Dictionary<Workbook, List<Window>>();
            this.excelApplication = excelApplication;
            var monitor = new WorkbookClosedMonitor(excelApplication);
            monitor.WorkbookClosed += MonitorWorkbookClosed;
        }

        void MonitorWorkbookClosed(object sender, WorkbookClosedEventArgs e)
        {
            VstoContribLog.Debug(log => log("Excel raised WorkbookClosed({0}) event", e.Workbook.ToLogFormat()));
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

            if (!excelApplication.Workbooks.OfType<Workbook>().Except(new[]{e.Workbook}).Any())
            {
                nullContextOpen = true;
                foreach (Window window in excelApplication.Windows)
                {
                    NewView(this, new NewViewEventArgs(window, NullContext.Instance, ExcelRibbonType.ExcelWorkbook.GetEnumDescription()));
                }
            }
        }

        /// <summary>
        /// Initialises this instance.
        /// </summary>
        public void Initialise()
        {
            ((AppEvents_Event)excelApplication).NewWorkbook += NewWorkbook;
            excelApplication.WorkbookOpen += WorkbookOpen;
        }

        void NewWorkbook(Workbook wb)
        {
            VstoContribLog.Debug(log => log("Excel raised NewWorkbook({0}) event", wb.ToLogFormat()));
            OnInitialise(wb);
        }

        void WorkbookOpen(Workbook wb)
        {
            VstoContribLog.Debug(log => log("Excel raised WorkbookOpen({0}) event", wb.ToLogFormat()));
            OnInitialise(wb);
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
                if (nullContextOpen)
                    ViewClosed(this, new ViewClosedEventArgs(singleWindow, NullContext.Instance));
                NewView(this, new NewViewEventArgs(singleWindow, wb, ExcelRibbonType.ExcelWorkbook.GetEnumDescription()));
            }
            else
            {
                foreach (Window window in wb.Windows)
                {
                    workbooks[wb].Add(window);
                    if (nullContextOpen)
                        ViewClosed(this, new ViewClosedEventArgs(window, NullContext.Instance));
                    NewView(this, new NewViewEventArgs(window, wb, ExcelRibbonType.ExcelWorkbook.GetEnumDescription()));
                }
            }

            nullContextOpen = false;
            wb.WindowActivate += wn => Activate(wb, wn);
        }

        void Activate(Workbook wb, Window wn)
        {
            VstoContribLog.Debug(log => log("Excel raised WorkbookOpen(wb: {0}, wn: {1}) event", wb.ToLogFormat(), wn.ToLogFormat()));
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
        }

        bool IsMdi()
        {
            return new Version(excelApplication.Version).Major <= 14;
        }

        public event EventHandler<NewViewEventArgs> NewView = (sender, args) => { };
        public event EventHandler<ViewClosedEventArgs> ViewClosed = (sender, args) => { };

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