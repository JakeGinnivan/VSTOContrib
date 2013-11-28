using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;

namespace VSTOContrib.Core.RibbonFactory
{
    /// <summary>
    ///     Because you cannot make a generic type COM visible, moving all code that requires generics into this class
    /// </summary>
    /// <typeparam name="TRibbonTypes"></typeparam>
    public class RibbonFactoryController<TRibbonTypes> : IRibbonFactoryController where TRibbonTypes : struct
    {
        readonly ViewModelResolver<TRibbonTypes> ribbonViewModelResolver;

        IViewProvider<TRibbonTypes> viewProvider;
        readonly VstoContribContext<TRibbonTypes> vstoContribContext;
        readonly RibbonXmlRewriter<TRibbonTypes> ribbonXmlRewriter;

        /// <summary>
        ///     Initializes a new instance of the <see cref="RibbonFactoryController{TRibbonTypes}" /> class.
        /// </summary>
        /// <param name="assemblies">The assemblies.</param>
        /// <param name="viewContextProvider">The view context provider</param>
        /// <param name="viewModelFactory">A delegate taking a type and returning an instance of the requested type</param>
        /// <param name="customTaskPaneCollection"></param>
        /// <param name="vstoFactory">The VSTO factory</param>
        /// <param name="viewLocationStrategy">The view location strategy.</param>
        public RibbonFactoryController(
            ICollection<Assembly> assemblies,
            IViewContextProvider viewContextProvider,
            IViewModelFactory viewModelFactory,
            Func<object> customTaskPaneCollection, 
            Factory vstoFactory, 
            IViewLocationStrategy viewLocationStrategy = null)
        {
            if (assemblies.Count == 0)
                throw new InvalidOperationException("You must specify at least one assembly to scan for viewmodels");

            vstoContribContext = new VstoContribContext<TRibbonTypes>();
            var ribbonTypes = GetTRibbonTypesInAssemblies(assemblies).ToList();

            ribbonViewModelResolver = new ViewModelResolver<TRibbonTypes>(
                ribbonTypes, new CustomTaskPaneRegister(customTaskPaneCollection), viewContextProvider, viewModelFactory, vstoFactory);

            ribbonXmlRewriter = new RibbonXmlRewriter<TRibbonTypes>(
                viewLocationStrategy ?? new DefaultViewLocationStrategy(),
                vstoContribContext, ribbonViewModelResolver);

            var loadExpression = ((Expression<Action<RibbonFactory>>)(r => r.Ribbon_Load(null)));
            string loadMethodName = loadExpression.GetMethodName();

            foreach (Type viewModelType in ribbonTypes)
            {
                ribbonXmlRewriter.LocateAndRegisterViewXml(viewModelType, loadMethodName);
            }
        }

        /// <summary>
        ///     Initialises the specified view provider.
        /// </summary>
        /// <typeparam name="TRibbonType">The type of the ribbon type.</typeparam>
        /// <param name="viewProvider">The view provider.</param>
        /// <returns></returns>
        public void Initialise<TRibbonType>(IViewProvider<TRibbonType> viewProvider)
        {
            this.viewProvider = (IViewProvider<TRibbonTypes>)viewProvider;

            ribbonViewModelResolver.Initialise(this.viewProvider);

            this.viewProvider.Initialise();
        }
        
        /// <summary>
        ///     Gets the custom UI.
        /// </summary>
        /// <param name="ribbonId">The ribbon id.</param>
        /// <returns></returns>
        public string GetCustomUI(string ribbonId)
        {
            TRibbonTypes enumFromDescription;
            try
            {
                enumFromDescription = EnumExtensions.EnumFromDescription<TRibbonTypes>(ribbonId);
            }
            catch (ArgumentException)
            {
                //An unknown ribbon type
                return null;
            }

            return !vstoContribContext.RibbonXmlFromTypeLookup.ContainsKey(enumFromDescription)
                       ? null
                       : vstoContribContext.RibbonXmlFromTypeLookup[enumFromDescription];
        }

        /// <summary>
        ///     Invokes the get.
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="caller">The caller.</param>
        /// <param name="parameters">The parameters.</param>
        /// <returns></returns>
        public object InvokeGet(IRibbonControl control, Expression<Action> caller, params object[] parameters)
        {
            CallbackTarget<TRibbonTypes> callbackTarget = vstoContribContext.TagToCallbackTargetLookup[control.Tag + caller.GetMethodName()];

            IRibbonViewModel viewModelInstance = ribbonViewModelResolver.ResolveInstanceFor(control.Context);

            Type type = viewModelInstance.GetType();
            PropertyInfo property = type.GetProperty(callbackTarget.Method);

            if (property != null)
            {
                return type.InvokeMember(callbackTarget.Method,
                                         BindingFlags.GetProperty,
                                         null,
                                         viewModelInstance,
                                         null);
            }

            try
            {
                return type.InvokeMember(callbackTarget.Method,
                                         BindingFlags.InvokeMethod,
                                         null,
                                         viewModelInstance,
                                         new[]
                                         {
                                             control
                                         }
                                             .Concat(parameters)
                                             .ToArray());
            }
            catch (MissingMethodException)
            {
                throw new InvalidOperationException(
                    string.Format("Expecting method with signature: {0}.{1}(IRibbonControl control)",
                                  type.Name,
                                  callbackTarget.Method));
            }
        }

        /// <summary>
        ///     Invokes the specified control.
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="caller">The caller.</param>
        /// <param name="parameters">The parameters.</param>
        public void Invoke(IRibbonControl control, Expression<Action> caller, params object[] parameters)
        {
            try
            {
                CallbackTarget<TRibbonTypes> callbackTarget = vstoContribContext.TagToCallbackTargetLookup[control.Tag + caller.GetMethodName()];

                IRibbonViewModel viewModelInstance = ribbonViewModelResolver.ResolveInstanceFor(control.Context);

                Type type = viewModelInstance.GetType();
                PropertyInfo property = type.GetProperty(callbackTarget.Method);

                if (property != null)
                {
                    type.InvokeMember(callbackTarget.Method,
                                      BindingFlags.SetProperty,
                                      null,
                                      viewModelInstance,
                                      new[]
                                      {
                                          parameters.Single()
                                      });
                }
                else
                {
                    type.InvokeMember(callbackTarget.Method,
                                      BindingFlags.InvokeMethod,
                                      null,
                                      viewModelInstance,
                                      new[]
                                      {
                                          control
                                      }
                                          .Concat(parameters)
                                          .ToArray());
                }
            }
            catch (Exception ex)
            {
                //TODO Provider better error handling, handle TargetInvocationException,
                // then surface inner exceptions message
                Debug.WriteLine(ex);
                throw;
            }
        }

        /// <summary>
        ///     Ribbons the loaded.
        /// </summary>
        /// <param name="ribbonUi">The ribbon UI.</param>
        public void RibbonLoaded(IRibbonUI ribbonUi)
        {
            ribbonViewModelResolver.RibbonLoaded(ribbonUi);
        }

        /// <summary>
        ///     Gets or sets the locate view strategy.
        /// </summary>
        /// <value>The locate view strategy.</value>
        public IViewLocationStrategy LocateViewStrategy
        {
            set { ribbonXmlRewriter.LocateViewStrategy = value; }
            get { return ribbonXmlRewriter.LocateViewStrategy; }
        }

        static IEnumerable<Type> GetTRibbonTypesInAssemblies(IEnumerable<Assembly> assemblies)
        {
            Type ribbonViewModelType = typeof(IRibbonViewModel);
            return assemblies
                .Select(
                    assembly =>
                    {
                        Type[] types = assembly.GetTypes();
                        return types.Where(ribbonViewModelType.IsAssignableFrom);
                    }
                )
                .Aggregate((t, t1) => t.Concat(t1));
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        /// <filterpriority>2</filterpriority>
        public void Dispose()
        {
            ribbonViewModelResolver.Dispose();
        }
    }
}