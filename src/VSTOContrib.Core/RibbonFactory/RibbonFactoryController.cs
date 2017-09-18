using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Microsoft.Office.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;

namespace VSTOContrib.Core.RibbonFactory
{
    /// <summary>
    ///     Because you cannot make a generic type COM visible, moving all code that requires generics into this class
    /// </summary>
    class RibbonFactoryController : IRibbonFactoryController
    {
        readonly ViewModelResolver ribbonViewModelResolver;
        readonly VstoContribContext vstoContribContext;
        readonly CustomTaskPaneRegister customTaskPaneRegister;
        IViewProvider viewProvider;

        public RibbonFactoryController(
            IViewContextProvider viewContextProvider,
            VstoContribContext vstoContribContext)
        {
            this.vstoContribContext = vstoContribContext;
            var ribbonTypes = GetTRibbonTypesInAssemblies(vstoContribContext.Assemblies).ToList();

            customTaskPaneRegister = new CustomTaskPaneRegister(vstoContribContext.AddinBase);
            ribbonViewModelResolver = new ViewModelResolver(
                ribbonTypes, customTaskPaneRegister, viewContextProvider, 
                vstoContribContext);

            var ribbonXmlRewriter = new RibbonXmlRewriter(vstoContribContext, ribbonViewModelResolver);

            var loadExpression = ((Expression<Action<RibbonFactory>>)(r => r.Ribbon_Load(null)));
            string loadMethodName = loadExpression.GetMethodName();

            foreach (Type viewModelType in ribbonTypes)
            {
                ribbonXmlRewriter.LocateAndRegisterViewXml(viewModelType, loadMethodName, vstoContribContext.FallbackRibbonType);
            }
        }

        public void Initialise(IViewProvider viewProvider)
        {
            this.viewProvider = viewProvider;

            ribbonViewModelResolver.Initialise(this.viewProvider);

            this.viewProvider.Initialise();
        }

        public string GetCustomUI(string ribbonId)
        {
            return !vstoContribContext.RibbonXmlFromTypeLookup.ContainsKey(ribbonId)
                       ? null
                       : vstoContribContext.RibbonXmlFromTypeLookup[ribbonId];
        }

        public string InvokeGetContent(IRibbonControl control, Expression<Action> caller, params object[] parameters)
        {
            // Remove any previous registered callbacks for this dynamic context
            vstoContribContext.RemoveCallbacksForDynamicContext(control.Tag);
            
            // Delegate to the view model to get the raw xml
            var xmlString = InvokeGet(control, caller, parameters);

            if (xmlString == null) return null;

            // Rewrite the XML with our callbacks, registering new callback targets
            var ribbonXmlRewriter = new RibbonXmlRewriter(vstoContribContext, ribbonViewModelResolver);
            var ribbonType = vstoContribContext.TagToCallbackTargetLookup[control.Tag + caller.GetMethodName()].RibbonType;
            return ribbonXmlRewriter.RewriteDynamicXml(ribbonType, control.Tag, xmlString.ToString());
        }

        public object InvokeGet(IRibbonControl control, Expression<Action> caller, params object[] parameters)
        {
            var methodName = caller.GetMethodName();
            CallbackTarget callbackTarget = vstoContribContext.TagToCallbackTargetLookup[control.Tag + methodName];

            var view = (object)control.Context;
            IRibbonViewModel viewModelInstance = ribbonViewModelResolver.ResolveInstanceFor(view);
            VstoContribLog.Debug(l => l("Ribbon callback {0} being invoked on {1} (View: {2}, ViewModel: {3})",
                methodName, control.Id, view.ToLogFormat(), viewModelInstance.ToLogFormat()));

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

        public void Invoke(IRibbonControl control, Expression<Action> caller, params object[] parameters)
        {
            try
            {
                var methodName = caller.GetMethodName();
                CallbackTarget callbackTarget =
                    vstoContribContext.TagToCallbackTargetLookup[control.Tag + methodName];

                var view = (object)control.Context;
                IRibbonViewModel viewModelInstance = ribbonViewModelResolver.ResolveInstanceFor(view);
                VstoContribLog.Debug(l => l("Ribbon callback {0} being invoked on {1} (View: {2}, ViewModel: {3})",
                    methodName, control.Id, view.ToLogFormat(), viewModelInstance.ToLogFormat()));

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
            catch (TargetInvocationException e)
            {
                var innerEx = e.InnerException;
                PreserveStackTrace(innerEx);
                if (vstoContribContext.ErrorHandlers != null && vstoContribContext.ErrorHandlers.Count == 0)
                {
                    Trace.TraceError(innerEx.ToString());
                }

                var handled = vstoContribContext.ErrorHandlers != null && vstoContribContext.ErrorHandlers.Any(errorHandler => errorHandler.Handle(innerEx));

                if (!handled)
                    throw innerEx;
            }
        }

        // http://weblogs.asp.net/fmarguerie/archive/2008/01/02/rethrowing-exceptions-and-preserving-the-full-call-stack-trace.aspx
        internal static void PreserveStackTrace(Exception exception)
        {
            MethodInfo preserveStackTrace = typeof(Exception).GetMethod("InternalPreserveStackTrace",
              BindingFlags.Instance | BindingFlags.NonPublic);
            preserveStackTrace.Invoke(exception, null);
        }

        public void RibbonLoaded(IRibbonUI ribbonUi)
        {
            ribbonViewModelResolver.RibbonLoaded(ribbonUi);
        }

        static IEnumerable<Type> GetTRibbonTypesInAssemblies(IEnumerable<Assembly> assemblies)
        {
            VstoContribLog.Debug(_ => _("Discovering ViewModels"));

            Type ribbonViewModelType = typeof(IRibbonViewModel);
            return assemblies
                .Select(assembly =>
                    {
                        VstoContribLog.Debug(_ => _("Discovering ViewModels in {0}", assembly.GetName().Name));
                        var types = assembly.GetTypes();
                        var viewModelTypes = types.Where(ribbonViewModelType.IsAssignableFrom).ToArray();
                        VstoContribLog.Debug(_ => _("Found:{0}", string.Join(string.Empty, viewModelTypes.Select(vm => "\r\n  " + vm.Name))));
                        return viewModelTypes;
                    }
                )
                .SelectMany(vm => vm);
        }

        public void Dispose()
        {
            ribbonViewModelResolver.Dispose();
            customTaskPaneRegister.Dispose();
        }
    }
}
