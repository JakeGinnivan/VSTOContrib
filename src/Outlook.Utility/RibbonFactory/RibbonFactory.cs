using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using Office.Utility;
using Action = System.Action;

namespace Outlook.Utility.RibbonFactory
{
    /// <summary>
    /// Simplifies adding custom Ribbon's to Office. 
    /// Allows the custom Ribbon xml to be wired up to IRibbonViewModel's
    /// by convention. Simply name the Ribbon.xml the same as the ribbon view model class
    /// in the same assembly
    /// </summary>
    [ComVisible(true)]
    public partial class RibbonFactory : IRibbonFactory
    {
        /// <summary>
        /// Lookup from a viewmodel type to it's ribbon XML
        /// </summary>
        private readonly Dictionary<RibbonType, string> _ribbonViews = new Dictionary<RibbonType, string>();
        private readonly Dictionary<string, CallbackTarget> _ribbonCallbackTarget = new Dictionary<string, CallbackTarget>();
        const string OfficeCustomui = "http://schemas.microsoft.com/office/2006/01/customui";
        const string OfficeCustomui4 = "http://schemas.microsoft.com/office/2009/07/customui";
        internal const string CommonCallbacks = "CommonCallbacks";
        private RibbonType _currentlyLoadingRibbon;
        private ControlCallbackLookup _controlCallbackLookup;
        private static IRibbonFactory _instance;
        private ViewLocationStrategyBase _viewLocationStrategy;
        
        private bool _initialsed;
        private ViewModelResolver _ribbonViewModelResolver;

        private RibbonFactory()
        {
            _viewLocationStrategy = new DefaultViewLocationStrategy();
        }

        /// <summary>
        /// Gets the instance of the ribbon factory.
        /// </summary>
        /// <value>The instance.</value>
        public static IRibbonFactory Instance
        {
            get { return _instance ?? (_instance = new RibbonFactory()); }
        }

        /// <summary>
        /// Initialises and builds up the ribbon factory
        /// </summary>
        /// <param name="ribbonFactory">The ribbon factory.</param>
        /// <param name="outlookApplication">The outlook application.</param>
        /// <param name="customTaskPaneCollection">The custom task pane collection.</param>
        /// <param name="assemblies">The assemblies to scan for view models.</param>
        /// <returns>
        /// Disposible object to call on outlook shutdown
        /// </returns>
        /// <exception cref="ViewNotFoundException">If the view cannot be located for a view model</exception>
        public IDisposable InitialiseFactory(Func<Type, IRibbonViewModel> ribbonFactory, Application outlookApplication, 
            CustomTaskPaneCollection customTaskPaneCollection, params Assembly[] assemblies)
        {
            if (_initialsed) throw new InvalidOperationException("Ribbon Factory already Initialised");
            _initialsed = true;

            var ribbonTypes = GetRibbonTypesInAssemblies(assemblies);

            _ribbonViewModelResolver = new ViewModelResolver(ribbonTypes, ribbonFactory, outlookApplication, customTaskPaneCollection);
            _controlCallbackLookup = new ControlCallbackLookup(GetRibbonElements());

            Expression<Action> loadMethod = () => Ribbon_Load(null);
            var loadMethodName = loadMethod.GetMethodName();


            foreach (var viewModelType in ribbonTypes)
            {
                LocateAndRegisterViewXml(viewModelType, loadMethodName);
            }

            return _ribbonViewModelResolver;
        }

        private static IEnumerable<Type> GetRibbonTypesInAssemblies(IEnumerable<Assembly> assemblies)
        {
            var ribbonViewModelType = typeof (IRibbonViewModel);
            return assemblies
                .Select(
                    assembly =>
                    assembly.GetTypes().Where(t => t.GetInterfaces().Any(ribbonViewModelType.IsAssignableFrom)))
                .Aggregate((t, t1) => t.Concat(t1));
        }

        private void LocateAndRegisterViewXml(Type viewModelType, string loadMethodName)
        {
            var resourceText = (string)_viewLocationStrategy.GetType()
                    .GetMethod("LocateViewForViewModel")
                    .MakeGenericMethod(viewModelType).
                    Invoke(_viewLocationStrategy, new object[] { });

            var ribbonDoc = XDocument.Parse(resourceText);

            //We have to override the Ribbon_Load event to make sure we get the callback
            var customUi = 
                ribbonDoc.Descendants(XName.Get("customUI", OfficeCustomui)).SingleOrDefault()
                ?? ribbonDoc.Descendants(XName.Get("customUI", OfficeCustomui4)).Single();

            customUi.SetAttributeValue("onLoad", loadMethodName);

            foreach (var value in RibbonViewModelHelper.GetRibbonTypesFor(viewModelType))
            {
                WireUpEvents(value, ribbonDoc, customUi.GetDefaultNamespace());
                _ribbonViews.Add(value, ribbonDoc.ToString());
            }
        }


        ///<summary>
        /// Gets or Sets the strategy that fetches the Ribbon XML for a given view
        ///</summary>
        public ViewLocationStrategyBase LocateViewStrategy
        {
            get { return _viewLocationStrategy; }
            set
            {
                if (value == null) return;

                _viewLocationStrategy = value;
            }
        }

        // ReSharper disable InconsistentNaming
        /// <summary>
        /// Ribbon_s the load.
        /// </summary>
        /// <param name="ribbonUi">The ribbon UI.</param>
        public void Ribbon_Load(IRibbonUI ribbonUi)
        {
            _ribbonViewModelResolver.RibbonLoaded(_currentlyLoadingRibbon, ribbonUi);
        }
        // ReSharper restore InconsistentNaming

        private void WireUpEvents(RibbonType ribbonType, XContainer ribbonDoc, XNamespace xNamespace)
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
                        var callbackTag = BuildTag(ribbonType, elementId, factoryMethodName);
                        _ribbonCallbackTarget.Add(callbackTag, new CallbackTarget(ribbonType, currentCallback));
                        xElement.SetAttributeValue(XName.Get("tag"), (ribbonType + elementId.Value));
                        _ribbonViewModelResolver.RegisterCallbackControl(ribbonType, currentCallback, elementId.Value);
                    }
                }
            }
        }

        private static string BuildTag(RibbonType viewModelType, XAttribute elementId, string factoryMethodName)
        {
            return viewModelType + elementId.Value + factoryMethodName;
        }

        /// <summary>
        /// Gets the custom UI.
        /// </summary>
        /// <param name="ribbonId">The ribbon id.</param>
        /// <returns></returns>
        public string GetCustomUI(string ribbonId)
        {
            RibbonType enumFromDescription;
            try
            {
                enumFromDescription = EnumExtensions.EnumFromDescription<RibbonType>(ribbonId);
            }
            catch (ArgumentException)
            {
                //An unknown ribbon type
                return null;
            }

            if (!_ribbonViews.ContainsKey(enumFromDescription)) return null;

            _currentlyLoadingRibbon = enumFromDescription;
            return _ribbonViews[enumFromDescription];
        }

        private object InvokeGet(IRibbonControl control, Expression<Action> caller, params object[] parameters)
        {
            var callbackTarget = _ribbonCallbackTarget[control.Tag + caller.GetMethodName()];

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

        private void Invoke(IRibbonControl control, Expression<Action> caller, params object[] parameters)
        {
            var callbackTarget = _ribbonCallbackTarget[control.Tag+caller.GetMethodName()];

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
    }
}
