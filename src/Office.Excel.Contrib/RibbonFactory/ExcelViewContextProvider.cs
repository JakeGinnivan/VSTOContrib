using Microsoft.Office.Interop.Excel;
using Office.Contrib.RibbonFactory.Interfaces;

namespace Office.Excel.Contrib.RibbonFactory
{
    /// <summary>
    /// Gets the document for a view
    /// </summary>
    public class ExcelViewContextProvider : IViewContextProvider
    {
        /// <summary>
        /// Gets the context for view.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <returns></returns>
        public object GetContextForView(object view)
        {
            //There seems to be no way get the current workbook for a window like every other office app
            // ThisWorkbook returns the workbook associated with any running code, so *should* be ok..

            var window = view as Window;
            if (window != null)
                return window.Application.ThisWorkbook;

            //var protectedWindow = view as ProtectedViewWindow;
            //if (protectedWindow != null)
            //    return protectedWindow.Document;

            return null;
        }

        /// <summary>
        /// Gets the ribbon type for view.
        /// </summary>
        /// <typeparam name="TRibbonType">The type of the ribbon type.</typeparam>
        /// <param name="view">The view.</param>
        /// <returns></returns>
        public TRibbonType GetRibbonTypeForView<TRibbonType>(object view)
        {
            return (TRibbonType)(object)ExcelRibbonType.ExcelWorkbook;
        }
    }
}