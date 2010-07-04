using System.Collections.Generic;
using Microsoft.Office.Core;

namespace Office.Utility
{
    /// <summary>
    /// Instance of a Ribbon Factory
    /// </summary>
    public interface IRibbonFactory : IRibbonExtensibility
    {
        /// <summary>
        /// Initialises and builds up the ribbon factory
        /// </summary>
        /// <param name="ribbons">Ribbon view models to wire up</param>
        /// <exception cref="ViewNotFoundException">If the view cannot be located for a view model</exception>
        void InitialiseFactory(IEnumerable<IRibbonViewModel> ribbons);
    }
}