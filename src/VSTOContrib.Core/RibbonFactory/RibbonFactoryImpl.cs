using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Xml.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;

namespace VSTOContrib.Core.RibbonFactory
{
    /// <summary>
    /// Because you cannot make a generic type COM visible, moving all code that requires generics into this class
    /// </summary>
    /// <typeparam name="TRibbonTypes"></typeparam>
    public class RibbonFactoryImpl<TRibbonTypes> : IRibbonFactoryImpl where TRibbonTypes : struct
    {
        private IViewLocationStrategy _viewLocationStrategy;

        const string OfficeCustomui = "http://schemas.microsoft.com/office/2006/01/customui";
        const string OfficeCustomui4 = "http://schemas.microsoft.com/office/2009/07/customui";

        /// <summary>
        /// Lookup from a viewmodel type to it's ribbon XML
        /// </summary>
        private readonly Dictionary<string, CallbackTarget<TRibbonTypes>> _tagToCallbackTargetLookup;
        private readonly Dictionary<TRibbonTypes, string> _ribbonXmlFromTypeLookup;
        private readonly ViewModelResolver<TRibbonTypes> _ribbonViewModelResolver;
        private readonly CustomTaskPaneRegister _customTaskPaneRegister;
        private readonly ControlCallbackLookup _controlCallbackLookup;
        private readonly RibbonViewModelHelper _ribbonViewModelHelper;
        private IViewProvider<TRibbonTypes> _viewProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="RibbonFactoryImpl&lt;TRibbonTypes&gt;"/> class.
        /// </summary>
        /// <param name="viewLocationStrategy">The view location strategy.</param>
        /// <param name="assemblies">The assemblies.</param>
        public RibbonFactoryImpl(
            ICollection<Assembly> assemblies,
            IViewLocationStrategy viewLocationStrategy = null)
        {
            if (assemblies.Count == 0)
                throw new InvalidOperationException("You must specify at least one assembly to scan for viewmodels");

            _controlCallbackLookup = new ControlCallbackLookup(GetRibbonElements());
            _ribbonXmlFromTypeLookup = new Dictionary<TRibbonTypes, string>();
            _ribbonViewModelHelper = new RibbonViewModelHelper();
            _tagToCallbackTargetLookup = new Dictionary<string, CallbackTarget<TRibbonTypes>>();
            _viewLocationStrategy = viewLocationStrategy ?? new DefaultViewLocationStrategy();
            var ribbonTypes = GetTRibbonTypesInAssemblies(assemblies).ToList();

            _customTaskPaneRegister = new CustomTaskPaneRegister();
            _ribbonViewModelResolver = new ViewModelResolver<TRibbonTypes>(
                ribbonTypes, _ribbonViewModelHelper, _customTaskPaneRegister);

            var loadExpression = ((Expression<Action<RibbonFactory>>)(r => r.Ribbon_Load(null)));
            var loadMethodName = loadExpression.GetMethodName();

            foreach (var viewModelType in ribbonTypes)
            {
                LocateAndRegisterViewXml(viewModelType, loadMethodName);
            }
        }

        /// <summary>
        /// Initialises the specified view provider.
        /// </summary>
        /// <typeparam name="TRibbonType">The type of the ribbon type.</typeparam>
        /// <param name="viewProvider">The view provider.</param>
        /// <param name="ribbonFactory">The ribbon factory.</param>
        /// <param name="viewContextProvider">The view context provider.</param>
        /// <param name="customTaskPaneCollection">The custom task pane collection.</param>
        /// <returns></returns>
        public IDisposable Initialise<TRibbonType>(
            IViewProvider<TRibbonType> viewProvider,
            Func<Type, IRibbonViewModel> ribbonFactory,
            IViewContextProvider viewContextProvider,
            CustomTaskPaneCollection customTaskPaneCollection)
        {
            _viewProvider = (IViewProvider<TRibbonTypes>)viewProvider;

            _ribbonViewModelResolver.Initialise(ribbonFactory, _viewProvider, viewContextProvider);

            _customTaskPaneRegister.Initialise(customTaskPaneCollection);

            _viewProvider.Initialise();

            return _ribbonViewModelResolver;
        }

        /// <summary>
        /// Locates the and register view XML.
        /// </summary>
        /// <param name="viewModelType">Type of the view model.</param>
        /// <param name="loadMethodName">Name of the load method.</param>
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

        /// <summary>
        /// Gets the custom UI.
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

            return !_ribbonXmlFromTypeLookup.ContainsKey(enumFromDescription)
                ? null
                : _ribbonXmlFromTypeLookup[enumFromDescription];
        }

        /// <summary>
        /// Invokes the get.
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="caller">The caller.</param>
        /// <param name="parameters">The parameters.</param>
        /// <returns></returns>
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

        /// <summary>
        /// Invokes the specified control.
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="caller">The caller.</param>
        /// <param name="parameters">The parameters.</param>
        public void Invoke(IRibbonControl control, Expression<Action> caller, params object[] parameters)
        {
            try
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
            catch (Exception ex)
            {
                //TODO Provider better error handling, handle TargetInvocationException,
                // then surface inner exceptions message
                Debug.WriteLine(ex);
                throw;
            }
            
        }

        /// <summary>
        /// Ribbons the loaded.
        /// </summary>
        /// <param name="ribbonUi">The ribbon UI.</param>
        public void RibbonLoaded(IRibbonUI ribbonUi)
        {
            _ribbonViewModelResolver.RibbonLoaded(ribbonUi);
        }

        /// <summary>
        /// Gets or sets the locate view strategy.
        /// </summary>
        /// <value>The locate view strategy.</value>
        public IViewLocationStrategy LocateViewStrategy
        {
            get { return _viewLocationStrategy; }
            set
            {
                if (value == null) return;
                _viewLocationStrategy = value;
            }
        }

        private static Dictionary<string, Dictionary<string, Expression<Action<RibbonFactory>>>> GetRibbonElements()
        {
            return new Dictionary<string, Dictionary<string, Expression<Action<RibbonFactory>>>>
                                  {
                                      {RibbonFactory.CommonCallbacks, new Dictionary<string, Expression<Action<RibbonFactory>>>
                                                            {
                                                                {"getVisible", f=>f.GetVisible(null)},
                                                                {"getSupertip", f=>f.GetSuperTip(null)},
                                                                {"getScreentip", f=>f.GetScreenTip(null)},
                                                                {"getSize", f=>f.GetSize(null)},
                                                                {"getKeytip", f=>f.GetKeyTip(null)},
                                                                {"getLabel", f=>f.GetLabel(null)},
                                                                {"getImageMso",f=>f.GetImageMso(null)},
                                                                {"getImage", f=>f.GetImage(null)},
                                                                {"getEnabled", f=>f.GetEnabled(null)},
                                                                {"getDescription", f=>f.GetDescription(null)}
                                                            }},
                                      {"button", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                                     {
                                                         {"onAction", f=>f.OnAction(null)},
                                                         {"getShowLabel", f=>f.GetShowLabel(null)},
                                                         {"getShowImage", f=>f.GetShowImage(null)}
                                                     }},
                                      {"checkBox", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                                       {
                                                           {"onAction", f=>f.PressedOnAction(null, true)},
                                                           {"getPressed", f=>f.GetPressed(null)}
                                                       }},
                                      {"dropDown", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                                       {
                                                           {"getItemCount", f=>f.GetItemCount(null)},
                                                           {"getItemID", f=>f.GetItemId(null, 0)},
                                                           {"getItemImage", f=>f.GetItemImage(null, 0)},
                                                           {"getItemLabel", f=>f.GetItemLabel(null, 0)},
                                                           {"getItemScreenTip", f=>f.GetItemScreenTip(null, 0)},
                                                           {"getItemSuperTip", f=>f.GetItemSuperTip(null, 0)},
                                                           {"getSelectedItemID", f=>f.GetSelectedItemId(null)},
                                                           {"getSelectedItemIndex", f=>f.GetSelectedItemIndex(null)},
                                                           {"onAction", f=>f.SelectionOnAction(null, null, 0)}
                                                       }},
                                      {"dynamicMenu", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                                          {
                                                              {"getContent", f=>f.GetContent(null)}
                                                          }},
                                      {"gallery", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                                      {
                                                          {"getItemCount", f=>f.GetItemCount(null)},
                                                          {"getItemHeight", f=>f.GetItemHeight(null)},
                                                          {"getItemID", f=>f.GetItemId(null, 0)},
                                                          {"getItemImage", f=>f.GetItemImage(null, 0)},
                                                          {"getItemLabel", f=>f.GetItemLabel(null, 0)},
                                                          {"getItemScreenTip", f=>f.GetItemScreenTip(null, 0)},
                                                          {"getItemSuperTip", f=>f.GetItemSuperTip(null, 0)},
                                                          {"getSelectedItemID", f=>f.GetSelectedItemId(null)},
                                                          {"getSelectedItemIndex", f=>f.GetSelectedItemIndex(null)},
                                                          {"onAction", f=>f.SelectionOnAction(null, null, 0)}
                                                      }},
                                      {"menuSeparator",  new Dictionary<string, Expression<Action<RibbonFactory>>>
                                                             {
                                                                 {"getTitle", f=>f.GetTitle(null)}
                                                             }},
                                      {"toggleButton", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                                           {
                                                               {"getPressed", f=>f.GetPressed(null)},
                                                               {"onAction", f=>f.PressedOnAction(null, true)}
                                                           }},
                                      {"comboBox", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                                       {
                                                           {"getItemCount", f=>f.GetItemCount(null)},
                                                           {"getItemID", f=>f.GetItemId(null, 0)},
                                                           {"getItemImage", f=>f.GetItemImage(null, 0)},
                                                           {"getItemLabel", f=>f.GetItemLabel(null, 0)},
                                                           {"getItemScreenTip", f=>f.GetItemScreenTip(null, 0)},
                                                           {"getItemSuperTip", f=>f.GetItemSuperTip(null, 0)},
                                                           {"getText", f=>f.GetText(null)},
                                                           {"onChange", f=>f.OnTextChanged(null, null)}
                                                       }},
                                      {"editBox", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                                      {
                                                          {"getText", f=>f.GetText(null)},
                                                          {"onChange", f=>f.OnTextChanged(null, null)}
                                                      }}
                                  };
        }
    }
}
