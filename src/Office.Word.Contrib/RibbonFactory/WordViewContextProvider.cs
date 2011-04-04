using Microsoft.Office.Interop.Word;
using Office.Contrib.RibbonFactory.Interfaces;

namespace Office.Word.Contrib.RibbonFactory
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
        public object GetContextForView(object view)
        {
            var window = view as Window;
            if (window != null)
                return window.Document;

            //var protectedWindow = view as ProtectedViewWindow;
            //if (protectedWindow != null)
            //    return protectedWindow.Document;

            return null;
        }

        public TRibbonType GetRibbonTypeForView<TRibbonType>(object view)
        {
            return (TRibbonType)(object)WordRibbonType.WordDocument;
        }
    }
}