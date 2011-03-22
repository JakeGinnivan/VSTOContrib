using Microsoft.Office.Interop.Word;
using Office.Contrib.RibbonFactory.Interfaces;

namespace Office.Word.Contrib.RibbonFactory
{
    /// <summary>
    /// Gets the document for a view
    /// </summary>
    public class WordViewContextProvider : IViewContextProvider
    {
        public object GetContextForView(object view)
        {
            var window = view as Window;
            if (window != null)
                return window.Document;

            var protectedWindow = view as ProtectedViewWindow;
            if (protectedWindow != null)
                return protectedWindow.Document;

            return null;
        }
    }
}