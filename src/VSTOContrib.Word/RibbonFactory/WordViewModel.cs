using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Word.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    [WordRibbonViewModel]
    public abstract class WordRibbonViewModel : IRibbonViewModel
    {
        /// <summary>
        /// The associated ribbon, you can invalidate controls getting them to refresh
        /// their state through the IRibbonUI.
        /// </summary>
        /// <value>The ribbon UI.</value>
        public IRibbonUI RibbonUi { get; set; }

        /// <summary>
        /// Used to get the VSTO versions of the context. For example
        /// 
        /// var vstodocument = ((Microsoft.Office.Tools.Word.ApplicationFactory)VstoFactory).GetVstoObject(document);
        /// </summary>
        public Factory VstoFactory { get; set; }

        /// <summary>
        /// Called when the window that the ribbon is shown in is opened
        /// </summary>
        /// <param name="context">The context.</param>
        public void Initialised(object context)
        {

        }

        /// <summary>
        /// Called when the current view is changed.
        /// </summary>
        /// <param name="currentView">The current view.</param>
        public void CurrentViewChanged(object currentView)
        {
            var window = currentView as Window;
            if (window != null)
                CurrentViewChanged(window);

            //Commented due to 2007 not supporting protected view window, will reintroduce later
            //var protectedWindow = currentView as ProtectedViewWindow;
            //if (protectedWindow != null)
            //    CurrentViewChanged(protectedWindow);
        }

        /// <summary>
        /// Called when the current view is changed.
        /// </summary>
        /// <param name="window">The window.</param>
        public abstract void CurrentViewChanged(Window window);

        ///// <summary>
        ///// Called when the current view is  changed.
        ///// </summary>
        ///// <param name="window">The window.</param>
        //public abstract void CurrentViewChanged(ProtectedViewWindow window);

        /// <summary>
        /// Cleanups this instance.
        /// </summary>
        public void Cleanup()
        {

        }
    }
}
