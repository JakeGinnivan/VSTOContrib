using System;
using Microsoft.Office.Interop.Excel;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;

namespace VSTOContrib.Excel.RibbonFactory
{
    public class ExcelOfficeApplicationEvents : IOfficeApplicationEvents
    {
        const string CaptionSuffix = " - Excel";
        const string ExcelLpClassName = "XLMAIN";
        Application excelApplication;

        void MonitorWorkbookClosed(object sender, WorkbookClosedEventArgs e)
        {
            VstoContribLog.Debug(log => log("Excel raised WorkbookClosed({0}) event", e.Workbook.ToLogFormat()));

            ContextClosed(e.Workbook);
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
            excelApplication.WindowActivate += Activate;
        }

        public event Action<NewViewEventArgs> NewView = _ => { };
        public event Action<OfficeWin32Window> ViewClosed = _ => { };
        public event Action<object> ContextClosed = _ => { };

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
            foreach (var window in wb.Windows)
            {
                var newViewEventArgs = new NewViewEventArgs(ToOfficeWindow(window), wb, ExcelRibbonType.ExcelWorkbook.GetEnumDescription());
                NewView(newViewEventArgs);
            }
        }

        void Activate(Workbook wb, Window wn)
        {
            var officeWin32Window = new OfficeWin32Window(wn, ExcelLpClassName, CaptionSuffix);
            VstoContribLog.Debug(log => log("Excel raised WorkbookOpen(wb: {0}, wn: {1}) event", wb.ToLogFormat(), wn.ToLogFormat()));
            var excelWorkbook = ExcelRibbonType.ExcelWorkbook.GetEnumDescription();
            NewView(new NewViewEventArgs(officeWin32Window, wb, excelWorkbook));
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
            foreach (Workbook wb in excelApplication.Workbooks)
            {
                OnInitialise(wb);
            }
        }
    }
}