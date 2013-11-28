using Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    /// <summary>
    /// Instance of a Ribbon Factory
    /// </summary>
    public interface IRibbonFactory : IRibbonExtensibility
    {
        /// <summary>
        /// Gets or sets the locate view strategy.
        /// </summary>
        /// <value>The locate view strategy.</value>
        IViewLocationStrategy LocateViewStrategy { get; set; }

        /// <summary>
        /// Sets the application the add-in is hosted in
        /// </summary>
        void SetApplication(object application, AddInBase addinBase);
    }
}