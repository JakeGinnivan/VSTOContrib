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
        readonly Dictionary<Workbook, List<OfficeWin32Window>> workbooks;
        const string CaptionSuffix = " - Excel";
        const string ExcelLpClassName = "XLMAIN";
        Application excelApplication;
        OfficeWin32Window singleWindow;
        bool nullContextOpen;

        public ExcelViewProvider()
        {
            workbooks = new Dictionary<Workbook, List<OfficeWin32Window>>();
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
                    NewView(this, new NewViewEventArgs(new OfficeWin32Window(window, ExcelLpClassName, CaptionSuffix), NullContext.Instance, ExcelRibbonType.ExcelWorkbook.GetEnumDescription()));
                }
            }
        }

        /// <summary>
        /// Initialises this instance.
        /// </summary>
        /// <param name="application"></param>
        public void Initialise(object application)
        {
            excelApplication = (Application)application;
            ((AppEvents_Event)excelApplication).NewWorkbook += NewWorkbook;
            excelApplication.WorkbookOpen += WorkbookOpen;
            var monitor = new WorkbookClosedMonitor(excelApplication);
            monitor.WorkbookClosed += MonitorWorkbookClosed;
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
                workbooks.Add(wb, new List<OfficeWin32Window>());

            if (IsMdi())
            {
                if (singleWindow == null)
                    singleWindow = new OfficeWin32Window(wb.Windows[1], ExcelLpClassName, CaptionSuffix);
                workbooks[wb].Add(singleWindow);
                if (nullContextOpen)
                    ViewClosed(this, new ViewClosedEventArgs(new OfficeWin32Window(singleWindow, ExcelLpClassName, CaptionSuffix), NullContext.Instance));
                NewView(this, new NewViewEventArgs(new OfficeWin32Window(singleWindow, ExcelLpClassName, CaptionSuffix), wb, ExcelRibbonType.ExcelWorkbook.GetEnumDescription()));
            }
            else
            {
                foreach (Window window in wb.Windows)
                {
                    var officeWin32Window = new OfficeWin32Window(window, ExcelLpClassName, CaptionSuffix);
                    workbooks[wb].Add(officeWin32Window);
                    if (nullContextOpen)
                        ViewClosed(this, new ViewClosedEventArgs(officeWin32Window, NullContext.Instance));
                    NewView(this, new NewViewEventArgs(officeWin32Window, wb, ExcelRibbonType.ExcelWorkbook.GetEnumDescription()));
                }
            }

            nullContextOpen = false;
            wb.WindowActivate += wn => Activate(wb, wn);
        }

        void Activate(Workbook wb, Window wn)
        {
            var officeWin32Window = new OfficeWin32Window(wn, ExcelLpClassName, CaptionSuffix);
            VstoContribLog.Debug(log => log("Excel raised WorkbookOpen(wb: {0}, wn: {1}) event", wb.ToLogFormat(), wn.ToLogFormat()));
            if (IsMdi() && !workbooks[wb].Contains(singleWindow))
                workbooks[wb].Add(singleWindow);
            if (!IsMdi() && !workbooks[wb].Contains(officeWin32Window))
            {
                var windows = workbooks[wb];
                if (windows.All(w => !w.Equals(officeWin32Window)))
                {
                    windows.Add(officeWin32Window);
                    var excelWorkbook = ExcelRibbonType.ExcelWorkbook.GetEnumDescription();
                    NewView(this, new NewViewEventArgs(officeWin32Window, wb, excelWorkbook));
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
        public void CleanupReferencesTo(OfficeWin32Window view, object context)
        {

        }

        public OfficeWin32Window ToOfficeWindow(object view)
        {
            return new OfficeWin32Window(view, ExcelLpClassName, CaptionSuffix);
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