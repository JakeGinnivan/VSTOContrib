using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Xml.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using Office.Contrib.RibbonFactory.Interfaces;

namespace Office.Contrib.RibbonFactory
{
    /// <summary>
    /// Because you cannot make a generic type COM visible, moving all code that requires generics into this class
    /// </summary>
    /// <typeparam name="TRibbonTypes"></typeparam>
    internal class RibbonFactoryImpl<TRibbonTypes> : IRibbonFactoryImpl where TRibbonTypes : struct
    {
        private readonly IViewProvider<TRibbonTypes> _viewProvider;
        private readonly IViewLocationStrategy _viewLocationStrategy;

        const string OfficeCustomui = "http://schemas.microsoft.com/office/2006/01/customui";
        const string OfficeCustomui4 = "http://schemas.microsoft.com/office/2009/07/customui";

        /// <summary>
        /// Lookup from a viewmodel type to it's ribbon XML
        /// </summary>
        private readonly Dictionary<TRibbonTypes, string> _ribbonXmlFromTypeLookup = new Dictionary<TRibbonTypes, string>();
        private readonly Dictionary<string, CallbackTarget<TRibbonTypes>> _tagToCallbackTargetLookup =
            new Dictionary<string, CallbackTarget<TRibbonTypes>>();
        private ViewModelResolver<TRibbonTypes> _ribbonViewModelResolver;
        private readonly RibbonViewModelHelper _ribbonViewModelHelper = new RibbonViewModelHelper();
        private ControlCallbackLookup _controlCallbackLookup;

        public RibbonFactoryImpl(
            IViewProvider<TRibbonTypes> viewProvider,
            IViewLocationStrategy viewLocationStrategy)
        {
            _viewProvider = viewProvider;
            _viewLocationStrategy = viewLocationStrategy;
        }

        public IDisposable Initialise(
            string loadMethodName,
            Dictionary<string, Dictionary<string, Expression<Action>>> ribbonElements,
            Func<Type, IRibbonViewModel> ribbonFactory,
            CustomTaskPaneCollection customTaskPaneCollection,
            params Assembly[] assemblies)
        {
            var ribbonTypes = GetTRibbonTypesInAssemblies(assemblies).ToList();

            _ribbonViewModelResolver = new ViewModelResolver<TRibbonTypes>(
                ribbonTypes, ribbonFactory, _ribbonViewModelHelper, customTaskPaneCollection, _viewProvider);
            _controlCallbackLookup = new ControlCallbackLookup(ribbonElements);

            foreach (var viewModelType in ribbonTypes)
            {
                LocateAndRegisterViewXml(viewModelType, loadMethodName);
            }

            _viewProvider.Initialise();

            return _ribbonViewModelResolver;
        }

        public void LocateAndRegisterViewXml(Type viewModelType, string loadMethodName)
        {
            var resourceText = (string)_viewLocationStrategy.GetType()
                    .GetMethod("LocateViewForViewModel")
                    .MakeGenericMethod(viewModelType)
                    .Invoke(_viewLocationStrategy, new object[] { });

            var ribbonDoc = XDocument.Parse(resourceText);

            //We have to override the Ribbon_Load event to make sure we get the callback
            var customUi =
                ribbonDoc.Descendants(XName.Get("customUI", OfficeCustomui)).SingleOrDefault()
                ?? ribbonDoc.Descendants(XName.Get("customUI", OfficeCustomui4)).Single();

            customUi.SetAttributeValue("onLoad", loadMethodName);

            foreach (var value in _ribbonViewModelHelper.GetRibbonTypesFor<TRibbonTypes>(viewModelType))
            {
                WireUpEvents(value, ribbonDoc, customUi.GetDefaultNamespace());
                _ribbonXmlFromTypeLookup.Add(value, ribbonDoc.ToString());
            }
        }

        private void WireUpEvents(TRibbonTypes ribbonTypes, XContainer ribbonDoc, XNamespace xNamespace)
        {
            //Go through each type of Ribbon 
            foreach (var ribbonControl in _controlCallbackLookup.RibbonControls)
            {
                //Get each instance of that control in the ribbon definition file
                var xElements = ribbonDoc.Descendants(XName.Get(ribbonControl, xNamespace.NamespaceName));

                foreach (var xElement in xElements)
                {
                    var elementId = xElement.Attribute(XName.Get("id"));
                    if (elementId == null) continue;

                    //Go through each possible callback, Concat with common methods on all controls
                    foreach (var controlCallback in _controlCallbackLookup.GetVstoControlCallbacks(ribbonControl))
                    {
                        //Look for a defined callback
                        var callbackAttribute = xElement.Attribute(XName.Get(controlCallback));

                        if (callbackAttribute == null) continue;
                        var currentCallback = callbackAttribute.Value;
                        //Set the callback value to the callback method defined on this factory
                        var factoryMethodName = _controlCallbackLookup.GetFactoryMethodName(ribbonControl, controlCallback);
                        callbackAttribute.SetValue(factoryMethodName);

                        //Set the tag attribute of the element, this is needed to know where to 
                        // direct the callback
                        var callbackTag = BuildTag(ribbonTypes, elementId, factoryMethodName);
                        _tagToCallbackTargetLookup.Add(callbackTag, new CallbackTarget<TRibbonTypes>(ribbonTypes, currentCallback));
                        xElement.SetAttributeValue(XName.Get("tag"), (ribbonTypes + elementId.Value));
                        _ribbonViewModelResolver.RegisterCallbackControl(ribbonTypes, currentCallback, elementId.Value);
                    }
                }
            }
        }

        private static string BuildTag(TRibbonTypes viewModelType, XAttribute elementId, string factoryMethodName)
        {
            return viewModelType + elementId.Value + factoryMethodName;
        }

        private static IEnumerable<Type> GetTRibbonTypesInAssemblies(IEnumerable<Assembly> assemblies)
        {
            var ribbonViewModelType = typeof(IRibbonViewModel);
            return assemblies
                .Select(
                    assembly =>
                    {
                        var types = assembly.GetTypes();
                        return types.Where(ribbonViewModelType.IsAssignableFrom);
                    }
                )
                .Aggregate((t, t1) => t.Concat(t1));
        }

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

            return !_ribbonXmlFromTypeLookup.ContainsKey(enumFromDescription)
                ? null
                : _ribbonXmlFromTypeLookup[enumFromDescription];
        }

        public object InvokeGet(IRibbonControl control, Expression<Action> caller, params object[] parameters)
        {
            var callbackTarget = _tagToCallbackTargetLookup[control.Tag + caller.GetMethodName()];

            var viewModelInstance = _ribbonViewModelResolver.ResolveInstanceFor(control.Context);

            Type type = viewModelInstance.GetType();
            var property = type.GetProperty(callbackTarget.Method);

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
            var callbackTarget = _tagToCallbackTargetLookup[control.Tag + caller.GetMethodName()];

            var viewModelInstance = _ribbonViewModelResolver.ResolveInstanceFor(control.Context);

            Type type = viewModelInstance.GetType();
            var property = type.GetProperty(callbackTarget.Method);

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

        public void RibbonLoaded(IRibbonUI ribbonUi)
        {
            _ribbonViewModelResolver.RibbonLoaded(ribbonUi);
        }
    }
}
