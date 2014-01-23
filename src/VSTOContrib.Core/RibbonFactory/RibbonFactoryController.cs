﻿using System;
using System.Collections.Generic;
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

        public object InvokeGet(IRibbonControl control, Expression<Action> caller, params object[] parameters)
        {
            CallbackTarget callbackTarget = vstoContribContext.TagToCallbackTargetLookup[control.Tag + caller.GetMethodName()];

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

        public void Invoke(IRibbonControl control, Expression<Action> caller, params object[] parameters)
        {
            try
            {
                CallbackTarget callbackTarget =
                    vstoContribContext.TagToCallbackTargetLookup[control.Tag + caller.GetMethodName()];

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
            catch (TargetInvocationException e)
            {
                var innerEx = e.InnerException;
                PreserveStackTrace(innerEx);
                var handled = false;
                for (int index = vstoContribContext.ErrorHandlers.Count - 1; index >= 0; index--)
                {
                    var handler = vstoContribContext.ErrorHandlers[index];
                    if (handler.Handle(innerEx))
                    {
                        handled = true;
                        break;
                    }
                }

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

        public void Dispose()
        {
            ribbonViewModelResolver.Dispose();
            customTaskPaneRegister.Dispose();
        }
    }
}