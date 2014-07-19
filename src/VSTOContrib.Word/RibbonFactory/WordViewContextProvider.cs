using Microsoft.Office.Interop.Word;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Word.RibbonFactory
{
    /// <summary>
    /// Gets the document for a view
    /// </summary>
    public class WordViewContextProvider : IViewContextProvider
    {
        /// <summary>
        /// Gets the context for view.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <returns></returns>
        public object GetContextForView(OfficeWin32Window view)
        {
            var window = view.Window as Window;
            if (window != null)
                return window.Document;

            //var protectedWindow = view as ProtectedViewWindow;
            //if (protectedWindow != null)
            //    return protectedWindow.Document;

            return null;
        }

        public string GetRibbonTypeForView(OfficeWin32Window view)
        {
            return WordRibbonType.WordDocument.GetEnumDescription();
        }
    }
}