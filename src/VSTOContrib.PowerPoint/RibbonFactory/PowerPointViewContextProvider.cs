using Microsoft.Office.Interop.PowerPoint;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.PowerPoint.RibbonFactory
{
    /// <summary>
    /// Gets the document for a view
    /// </summary>
    public class PowerPointViewContextProvider : IViewContextProvider
    {
        public object GetContextForView(OfficeWin32Window view)
        {
            var window = view.Window as DocumentWindow;
            if (window != null)
                return window.Presentation;

            //var slide = view as SlideShowView;
            //if (slide != null)
            //    return slide.Slide;

            //var slideWindow = view as SlideShowWindow;
            //if (slideWindow != null)
            //    return slideWindow.Presentation;

            //var protectedWindow = view as ProtectedViewWindow;
            //if (protectedWindow != null)
            //    return protectedWindow.Presentation;

            return null;
        }

        public string GetRibbonTypeForView(OfficeWin32Window view)
        {
            return PowerPointRibbonType.PowerPointPresentation.GetEnumDescription();
        }
    }
}