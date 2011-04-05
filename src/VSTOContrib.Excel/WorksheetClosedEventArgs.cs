using System;
using Microsoft.Office.Interop.Excel;

namespace VSTOContrib.Excel
{
    /// <summary>
    /// 
    /// </summary>
    public class WorkbookClosedEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookClosedEventArgs"/> class.
        /// </summary>
        /// <param name="workbook">The document.</param>
        public WorkbookClosedEventArgs(Workbook workbook)
        {
            Workbook = workbook;
        }

        /// <summary>
        /// Gets or sets the document.
        /// </summary>
        /// <value>The document.</value>
        public Workbook Workbook { get; set; }
    }
}