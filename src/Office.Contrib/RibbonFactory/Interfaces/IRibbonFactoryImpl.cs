using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace Office.Contrib.RibbonFactory.Interfaces
{
    /// <summary>
    /// 
    /// </summary>
    public interface IRibbonFactoryImpl
    {
        /// <summary>
        /// Initialises the specified view provider.
        /// </summary>
        /// <typeparam name="TRibbonTypes">The type of the ribbon types.</typeparam>
        /// <param name="viewProvider">The view provider.</param>
        /// <param name="loadMethodName">Name of the load method.</param>
        /// <param name="ribbonElements">The ribbon elements.</param>
        /// <param name="ribbonFactory">The ribbon factory.</param>
        /// <param name="customTaskPaneCollection">The custom task pane collection.</param>
        /// <param name="assemblies">The assemblies.</param>
        /// <returns></returns>
        IDisposable Initialise<TRibbonTypes>(
            IViewProvider<TRibbonTypes> viewProvider,
            string loadMethodName,
            Dictionary<string, Dictionary<string, Expression<Action>>> ribbonElements,
            Func<Type, IRibbonViewModel> ribbonFactory,
            CustomTaskPaneCollection customTaskPaneCollection,
            params Assembly[] assemblies);

        /// <summary>
        /// Gets the custom UI.
        /// </summary>
        /// <param name="ribbonId">The ribbon id.</param>
        /// <returns></returns>
        string GetCustomUI(string ribbonId);

        /// <summary>
        /// Invokes the get.
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="caller">The caller.</param>
        /// <param name="parameters">The parameters.</param>
        /// <returns></returns>
        object InvokeGet(IRibbonControl control, Expression<Action> caller, params object[] parameters);

        /// <summary>
        /// Invokes the specified control.
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="caller">The caller.</param>
        /// <param name="parameters">The parameters.</param>
        void Invoke(IRibbonControl control, Expression<Action> caller, params object[] parameters);

        /// <summary>
        /// Ribbons the loaded.
        /// </summary>
        /// <param name="ribbonUi">The ribbon UI.</param>
        void RibbonLoaded(IRibbonUI ribbonUi);

        /// <summary>
        /// Gets or sets the locate view strategy.
        /// </summary>
        /// <value>The locate view strategy.</value>
        IViewLocationStrategy LocateViewStrategy { get; set; }
    }
}