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
        /// The last visible or active window, explorer or inspector (depending on the Office application your are extending)
        /// </summary>
        object CurrentView { get; set; }

        /// <summary>
        /// Called when the window that the ribbon is shown in is opened
        /// </summary>
        /// <param name="context">The context.</param>
        void Initialised(object context);

        /// <summary>
        /// Cleanups this instance.
        /// </summary>
        void Cleanup();
    }
}
