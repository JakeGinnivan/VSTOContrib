using Microsoft.Office.Interop.Excel;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Excel.RibbonFactory
{
    /// <summary>
    /// Gets the document for a view
    /// </summary>
    public class ExcelViewContextProvider : IViewContextProvider
    {
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

        public string GetRibbonTypeForView(object view)
        {
            return ExcelRibbonType.ExcelWorkbook.GetEnumDescription();
        }
    }
}