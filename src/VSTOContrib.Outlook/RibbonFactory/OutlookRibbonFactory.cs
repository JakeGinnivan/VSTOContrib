using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Outlook.RibbonFactory
{
    /// <summary>
    /// Simplifies adding custom Ribbon's to Office. 
    /// Allows the custom Ribbon xml to be wired up to IRibbonViewModel's
    /// by convention. Simply name the Ribbon.xml the same as the ribbon view model class
    /// in the same assembly
    /// </summary>
    [ComVisible(true)]
    public class OutlookRibbonFactory : Core.RibbonFactory.RibbonFactory
    {
        readonly OutlookViewProvider viewProvider;

        public OutlookRibbonFactory(AddInBase addinBase, params Assembly[] assemblies):
            this(new OutlookViewProvider(), addinBase, UseIfEmpty(assemblies, Assembly.GetCallingAssembly()))
        {
        }

        private OutlookRibbonFactory(OutlookViewProvider viewProvider, AddInBase addinBase, Assembly[] assemblies)
            : base(addinBase, assemblies,
            new OutlookViewContextProvider(viewProvider), viewProvider, null)
        {
            this.viewProvider = viewProvider;
        }

        /// <summary>
        /// Initialisation callback for ribbon factory. The implementation must initialise the controller and 
        /// </summary>
        protected override void InitialiseRibbonFactoryController(IRibbonFactoryController controller, object application)
        {
            viewProvider.Initialise(application);
        }

        /// <summary>
        /// Called when the add-in is shutting down
        /// </summary>
        protected override void ShuttingDown()
        {
            viewProvider.Dispose();
        }
    }
}