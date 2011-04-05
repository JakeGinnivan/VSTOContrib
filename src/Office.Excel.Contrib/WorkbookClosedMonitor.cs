using System;
using Microsoft.Office.Interop.Excel;

namespace Office.Excel.Contrib
{
    /// <summary>
    /// Monitors for when a Excel workbook is closed
    /// </summary>
    public class WorkbookClosedMonitor
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookClosedMonitor"/> class.
        /// </summary>
        /// <param name="application">The application.</param>
        public WorkbookClosedMonitor(Application application)
        {
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }

            Application = application;

            Application.WorkbookActivate += ApplicationWorkbookActivate;
            Application.WorkbookBeforeClose += ApplicationWorkbookBeforeClose;
            Application.WorkbookDeactivate += ApplicationWorkbookDeactivate;
        }

        /// <summary>
        /// Occurs when workbook is closed.
        /// </summary>
        public event EventHandler<WorkbookClosedEventArgs> WorkbookClosed;

        /// <summary>
        /// Gets the application.
        /// </summary>
        /// <value>The application.</value>
        public Application Application { get; private set; }

        private CloseRequestInfo PendingRequest { get; set; }

        private void ApplicationWorkbookDeactivate(Workbook wb)
        {
            if (Application.Workbooks.Count != 1) return;

            // With only one workbook available deactivating means it will be closed
            PendingRequest = null;

            OnWorkbookClosed(new WorkbookClosedEventArgs(wb));
        }

        private void ApplicationWorkbookBeforeClose(Workbook wb, ref bool cancel)
        {
            if (!cancel)
            {
                PendingRequest = new CloseRequestInfo(wb, Application.Workbooks.Count);
            }
        }

        private void ApplicationWorkbookActivate(Workbook wb)
        {
            // A workbook was closed if a request is pending and the workbook count decreased
            var wasWorkbookClosed = PendingRequest != null
                                    && Application.Workbooks.Count < PendingRequest.WorkbookCount;

            if (wasWorkbookClosed)
            {
                var args = new WorkbookClosedEventArgs(PendingRequest.Workbook);

                PendingRequest = null;

                OnWorkbookClosed(args);
            }
            else
            {
                PendingRequest = null;
            }
        }

        private void OnWorkbookClosed(WorkbookClosedEventArgs e)
        {
            var handler = WorkbookClosed;

            if (handler != null)
                handler(this, e);
        }
    }
}