using Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    /// <summary>
    /// Ribbon ViewModel
    /// </summary>
    public interface IRibbonViewModel
    {
        /// <summary>
        /// The associated ribbon, you can invalidate controls getting them to refresh
        /// their state through the IRibbonUI.
        /// </summary>
        /// <value>The ribbon UI.</value>
        IRibbonUI RibbonUi { get; set; }

        /// <summary>
        /// Used to get the VSTO versions of the context. For example
        /// 
        /// var vstodocument = ((Microsoft.Office.Tools.Word.ApplicationFactory)VstoFactory).GetVstoObject(document);
        /// </summary>
        Factory VstoFactory { get; set; }

        /// <summary>
        /// Called when the window that the ribbon is shown in is opened
        /// </summary>
        /// <param name="context">The context.</param>
        void Initialised(object context);

        /// <summary>
        /// Called when the current view is changed.
        /// </summary>
        /// <param name="currentView">The current view.</param>
        void CurrentViewChanged(object currentView);

        /// <summary>
        /// Cleanups this instance.
        /// </summary>
        void Cleanup();
    }
}
