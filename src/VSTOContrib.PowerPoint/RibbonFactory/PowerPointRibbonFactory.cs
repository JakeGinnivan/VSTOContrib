using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.PowerPoint.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    [ComVisible(true)]
    public class PowerPointRibbonFactory : Core.RibbonFactory.RibbonFactory
    {
        private static Application powerPointApplication;
        private PowerPointViewProvider powerPointViewProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="PowerPointRibbonFactory"/> class.
        /// </summary>
        /// <param name="ribbonFactory">A delegate taking a type and returning an instance of the requested type</param>
        /// <param name="customTaskPaneCollection">A delayed resolution instance of the custom task pane collection of your addin 'new Lazy(()=>CustomTaskPaneCollection)'</param>
        /// <param name="assemblies">Assemblies to scan for view models</param>
        public PowerPointRibbonFactory(Func<Type, IRibbonViewModel> ribbonFactory,
            Lazy<CustomTaskPaneCollection> customTaskPaneCollection, params Assembly[] assemblies)
            : base(new RibbonFactoryController<PowerPointRibbonType>(assemblies, new PowerPointViewContextProvider(), ribbonFactory, customTaskPaneCollection))
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PowerPointRibbonFactory"/> class.
        /// </summary>
        /// <param name="ribbonFactory">A delegate taking a type and returning an instance of the requested type</param>
        /// <param name="customTaskPaneCollection">A delayed resolution instance of the custom task pane collection of your addin 'new Lazy(()=>CustomTaskPaneCollection)'</param>
        /// <param name="assemblies">Assemblies to scan for view models</param>
        /// <param name="viewLocationStrategy">The view location strategy, null for default strategy.</param>
        public PowerPointRibbonFactory(
            Func<Type, IRibbonViewModel> ribbonFactory,
            Lazy<CustomTaskPaneCollection> customTaskPaneCollection,
            IViewLocationStrategy viewLocationStrategy,
            params Assembly[] assemblies)
            : base(new RibbonFactoryController<PowerPointRibbonType>(assemblies, new PowerPointViewContextProvider(), ribbonFactory, customTaskPaneCollection, viewLocationStrategy))
        {
        }

        /// <summary>
        /// Initialises the ribbon factory.
        /// </summary>
        public override IDisposable InitialiseFactory(
            CustomTaskPaneCollection customTaskPaneCollection)
        {
            if (powerPointApplication == null)
                throw new InvalidOperationException("Set Word application instance first trough SetApplication()");

            powerPointViewProvider = new PowerPointViewProvider(powerPointApplication);
            return InitialiseFactoryInternal(
                powerPointViewProvider);
        }

        public override void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUi)
        {
            //Word does not raise a new document event when we are starting up, and initialise is too soon
            powerPointViewProvider.RegisterOpenDocuments();
            base.Ribbon_Load(ribbonUi);
        }

        /// <summary>
        /// Sets the PowerPoint application Instance
        /// </summary>
        /// <param name="application"></param>
        public static void SetApplication(Application application)
        {
            powerPointApplication = application;
        }
    }
}