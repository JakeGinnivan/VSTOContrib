using Microsoft.Office.Interop.Excel;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Excel.RibbonFactory
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
            var window = view as Window;
            if (window != null)
                return window.Application.ActiveWorkbook;

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