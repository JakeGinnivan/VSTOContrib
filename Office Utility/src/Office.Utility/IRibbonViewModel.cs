using Microsoft.Office.Core;

namespace Office.Utility
{
    /// <summary>
    /// Ribbon ViewModel
    /// </summary>
    public interface IRibbonViewModel
    {
        /// <summary>
        /// The type of Inspector or Explorer that the ribbon should be displayed for.
        /// </summary>
        /// <value>The ribbon type.</value>
        RibbonType Type { get; }

        /// <summary>
        /// The associated ribbon, you can invalidate controls getting them to refresh
        /// their state through the IRibbonUI.
        /// </summary>
        /// <value>The ribbon UI.</value>
        IRibbonUI RibbonUi { get; set; }
    }
}
