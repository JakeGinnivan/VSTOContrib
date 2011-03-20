using System;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using Office.Contrib.RibbonFactory.Interfaces;

namespace Office.Contrib.RibbonFactory
{
    /// <summary>
    /// Simplifies adding custom Ribbon's to Office. 
    /// Allows the custom Ribbon xml to be wired up to IRibbonViewModel's
    /// by convention. Simply name the Ribbon.xml the same as the ribbon view model class
    /// in the same assembly
    /// </summary>
    [ComVisible(true)]
    public abstract partial class RibbonFactory : IRibbonFactory
    {
        
        internal const string CommonCallbacks = "CommonCallbacks";
        
        private bool _initialsed;
        private static readonly object InstanceLock = new object();
        private readonly IRibbonFactoryImpl _ribbonFactoryImpl;

        /// <summary>
        /// Initializes a new instance of the <see cref="RibbonFactory"/> class.
        /// </summary>
        /// <param name="ribbonFactoryImpl"></param>
        protected RibbonFactory(IRibbonFactoryImpl ribbonFactoryImpl)
        {
            lock (InstanceLock)
            {
                if (Current != null)
                    throw new InvalidOperationException("You can only create a single ribbon factory");
                Current = this;
            }

            _ribbonFactoryImpl = ribbonFactoryImpl;
        }

        /// <summary>
        /// Initialises and builds up the ribbon factory
        /// </summary>
        /// <param name="ribbonFactory">The ribbon factory.</param>
        /// <param name="customTaskPaneCollection">The custom task pane collection.</param>
        /// <param name="assemblies">The assemblies to scan for view models.</param>
        /// <returns>
        /// Disposible object to call on outlook shutdown
        /// </returns>
        /// <exception cref="ViewNotFoundException">If the view cannot be located for a view model</exception>
        public abstract IDisposable InitialiseFactory(
            Func<Type, IRibbonViewModel> ribbonFactory,
            CustomTaskPaneCollection customTaskPaneCollection,
            params Assembly[] assemblies);

        /// <summary>
        /// Initialises the factory internal.
        /// </summary>
        /// <typeparam name="TRibbonTypes">The type of the ribbon types.</typeparam>
        /// <param name="viewProvider">The view provider.</param>
        /// <param name="ribbonFactory">The ribbon factory.</param>
        /// <param name="customTaskPaneCollection">The custom task pane collection.</param>
        /// <param name="assemblies">The assemblies.</param>
        /// <returns></returns>
        protected IDisposable InitialiseFactoryInternal<TRibbonTypes>(
            IViewProvider<TRibbonTypes> viewProvider,
            Func<Type, IRibbonViewModel> ribbonFactory,
            CustomTaskPaneCollection customTaskPaneCollection,
            params Assembly[] assemblies) where TRibbonTypes : struct
        {
            if (assemblies.Length == 0) 
                throw new InvalidOperationException("You must specify at least one assembly to scan for viewmodels");
            if (_initialsed)
                throw new InvalidOperationException("Ribbon Factory already Initialised");

            _initialsed = true;

            Expression<Action> loadMethod = () => Ribbon_Load(null);
            var loadMethodName = loadMethod.GetMethodName();

            return _ribbonFactoryImpl.Initialise(
                viewProvider, 
                loadMethodName, 
                GetRibbonElements(), 
                ribbonFactory, 
                customTaskPaneCollection,
                assemblies);
        }

        ///<summary>
        /// Gets or Sets the strategy that fetches the Ribbon XML for a given view
        ///</summary>
        public IViewLocationStrategy LocateViewStrategy
        {
            get { return _ribbonFactoryImpl.LocateViewStrategy; }
            set
            {
                if (value == null) return;

                _ribbonFactoryImpl.LocateViewStrategy = value;
            }
        }

        /// <summary>
        /// Current instance of RibbonFactory
        /// </summary>
        public static IRibbonFactory Current { get; protected set; }

        /// <summary>
        /// Ribbon_s the load.
        /// </summary>
        /// <param name="ribbonUi">The ribbon UI.</param>
        // ReSharper disable InconsistentNaming
        public void Ribbon_Load(IRibbonUI ribbonUi)
        {
            _ribbonFactoryImpl.RibbonLoaded(ribbonUi);
        }
        // ReSharper restore InconsistentNaming

        /// <summary>
        /// Gets the custom UI.
        /// </summary>
        /// <param name="ribbonId">The ribbon id.</param>
        /// <returns></returns>
        public string GetCustomUI(string ribbonId)
        {
            return _ribbonFactoryImpl.GetCustomUI(ribbonId);
        }
    }
}
