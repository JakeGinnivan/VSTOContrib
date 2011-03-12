using Microsoft.Office.Core;

namespace Office.Contrib.RibbonFactory
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
        /// Called when the window that the ribbon is shown in is opened
        /// </summary>
        /// <param name="context">The context.</param>
        void Displayed(object context);

        /// <summary>
        /// Cleanups this instance.
        /// </summary>
        void Cleanup();
    }
}
