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
    ///     Because you cannot make a generic type COM visible, moving all code that requires generics into this class
    /// </summary>
    /// <typeparam name="TRibbonTypes"></typeparam>
    public class RibbonFactoryController<TRibbonTypes> : IRibbonFactoryController where TRibbonTypes : struct
    {
        const string OfficeCustomui = "http://schemas.microsoft.com/office/2006/01/customui";
        const string OfficeCustomui4 = "http://schemas.microsoft.com/office/2009/07/customui";

        readonly ControlCallbackLookup controlCallbackLookup;
        readonly RibbonViewModelHelper ribbonViewModelHelper;
        readonly ViewModelResolver<TRibbonTypes> ribbonViewModelResolver;
        readonly Dictionary<TRibbonTypes, string> ribbonXmlFromTypeLookup;

        /// <summary>
        ///     Lookup from a viewmodel type to it's ribbon XML
        /// </summary>
        readonly Dictionary<string, CallbackTarget<TRibbonTypes>> tagToCallbackTargetLookup;

        IViewLocationStrategy viewLocationStrategy;

        IViewProvider<TRibbonTypes> viewProvider;

        /// <summary>
        ///     Initializes a new instance of the <see cref="RibbonFactoryController{TRibbonTypes}" /> class.
        /// </summary>
        /// <param name="assemblies">The assemblies.</param>
        /// <param name="viewContextProvider">The view context provider</param>
        /// <param name="ribbonFactory">A delegate taking a type and returning an instance of the requested type</param>
        /// <param name="customTaskPaneCollection"></param>
        /// <param name="viewLocationStrategy">The view location strategy.</param>
        public RibbonFactoryController(
            ICollection<Assembly> assemblies,
            IViewContextProvider viewContextProvider, 
            Func<Type, IRibbonViewModel> ribbonFactory, 
            Lazy<CustomTaskPaneCollection> customTaskPaneCollection, 
            IViewLocationStrategy viewLocationStrategy = null)
        {
            if (assemblies.Count == 0)
                throw new InvalidOperationException("You must specify at least one assembly to scan for viewmodels");

            controlCallbackLookup = new ControlCallbackLookup(GetRibbonElements());
            ribbonXmlFromTypeLookup = new Dictionary<TRibbonTypes, string>();
            ribbonViewModelHelper = new RibbonViewModelHelper();
            tagToCallbackTargetLookup = new Dictionary<string, CallbackTarget<TRibbonTypes>>();
            this.viewLocationStrategy = viewLocationStrategy ?? new DefaultViewLocationStrategy();
            List<Type> ribbonTypes = GetTRibbonTypesInAssemblies(assemblies).ToList();

            ribbonViewModelResolver = new ViewModelResolver<TRibbonTypes>(
                ribbonTypes, ribbonViewModelHelper, new CustomTaskPaneRegister(customTaskPaneCollection), viewContextProvider, ribbonFactory);

            var loadExpression = ((Expression<Action<RibbonFactory>>)(r => r.Ribbon_Load(null)));
            string loadMethodName = loadExpression.GetMethodName();

            foreach (Type viewModelType in ribbonTypes)
            {
                LocateAndRegisterViewXml(viewModelType, loadMethodName);
            }
        }

        /// <summary>
        ///     Initialises the specified view provider.
        /// </summary>
        /// <typeparam name="TRibbonType">The type of the ribbon type.</typeparam>
        /// <param name="viewProvider">The view provider.</param>
        /// <returns></returns>
        public IDisposable Initialise<TRibbonType>(
            IViewProvider<TRibbonType> viewProvider)
        {
            this.viewProvider = (IViewProvider<TRibbonTypes>)viewProvider;

            ribbonViewModelResolver.Initialise(this.viewProvider);

            this.viewProvider.Initialise();

            return ribbonViewModelResolver;
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

            return !ribbonXmlFromTypeLookup.ContainsKey(enumFromDescription)
                       ? null
                       : ribbonXmlFromTypeLookup[enumFromDescription];
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
            CallbackTarget<TRibbonTypes> callbackTarget =
                tagToCallbackTargetLookup[control.Tag + caller.GetMethodName()];

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
                CallbackTarget<TRibbonTypes> callbackTarget =
                    tagToCallbackTargetLookup[control.Tag + caller.GetMethodName()];

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
            get { return viewLocationStrategy; }
            set
            {
                if (value == null) return;
                viewLocationStrategy = value;
            }
        }

        /// <summary>
        ///     Locates the and register view XML.
        /// </summary>
        /// <param name="viewModelType">Type of the view model.</param>
        /// <param name="loadMethodName">Name of the load method.</param>
        public void LocateAndRegisterViewXml(Type viewModelType, string loadMethodName)
        {
            var resourceText = (string)viewLocationStrategy.GetType()
                                                             .GetMethod("LocateViewForViewModel")
                                                             .MakeGenericMethod(viewModelType)
                                                             .Invoke(viewLocationStrategy, new object[] { });

            XDocument ribbonDoc = XDocument.Parse(resourceText);

            //We have to override the Ribbon_Load event to make sure we get the callback
            XElement customUi =
                ribbonDoc.Descendants(XName.Get("customUI", OfficeCustomui)).SingleOrDefault()
                ?? ribbonDoc.Descendants(XName.Get("customUI", OfficeCustomui4)).Single();

            customUi.SetAttributeValue("onLoad", loadMethodName);

            //And for automatic image loading support
            if (customUi.Attribute("loadImage") == null)
                customUi.SetAttributeValue("loadImage", "GetPicture");

            foreach (TRibbonTypes value in ribbonViewModelHelper.GetRibbonTypesFor<TRibbonTypes>(viewModelType))
            {
                WireUpEvents(value, ribbonDoc, customUi.GetDefaultNamespace());
                ribbonXmlFromTypeLookup.Add(value, ribbonDoc.ToString());
            }
        }

        void WireUpEvents(TRibbonTypes ribbonTypes, XContainer ribbonDoc, XNamespace xNamespace)
        {
            //Go through each type of Ribbon 
            foreach (string ribbonControl in controlCallbackLookup.RibbonControls)
            {
                //Get each instance of that control in the ribbon definition file
                IEnumerable<XElement> xElements =
                    ribbonDoc.Descendants(XName.Get(ribbonControl, xNamespace.NamespaceName));

                foreach (XElement xElement in xElements)
                {
                    XAttribute elementId = xElement.Attribute(XName.Get("id"));
                    if (elementId == null) continue;

                    //Go through each possible callback, Concat with common methods on all controls
                    foreach (string controlCallback in controlCallbackLookup.GetVstoControlCallbacks(ribbonControl))
                    {
                        //Look for a defined callback
                        XAttribute callbackAttribute = xElement.Attribute(XName.Get(controlCallback));

                        if (callbackAttribute == null) continue;
                        string currentCallback = callbackAttribute.Value;
                        //Set the callback value to the callback method defined on this factory
                        string factoryMethodName = controlCallbackLookup.GetFactoryMethodName(ribbonControl,
                                                                                               controlCallback);
                        callbackAttribute.SetValue(factoryMethodName);

                        //Set the tag attribute of the element, this is needed to know where to 
                        // direct the callback
                        string callbackTag = BuildTag(ribbonTypes, elementId, factoryMethodName);
                        tagToCallbackTargetLookup.Add(callbackTag,
                                                       new CallbackTarget<TRibbonTypes>(ribbonTypes, currentCallback));
                        xElement.SetAttributeValue(XName.Get("tag"), (ribbonTypes + elementId.Value));
                        ribbonViewModelResolver.RegisterCallbackControl(ribbonTypes, currentCallback, elementId.Value);
                    }
                }
            }
        }

        static string BuildTag(TRibbonTypes viewModelType, XAttribute elementId, string factoryMethodName)
        {
            return viewModelType + elementId.Value + factoryMethodName;
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

        static Dictionary<string, Dictionary<string, Expression<Action<RibbonFactory>>>> GetRibbonElements()
        {
            return new Dictionary<string, Dictionary<string, Expression<Action<RibbonFactory>>>>
                   {
                       {
                           RibbonFactory.CommonCallbacks, new Dictionary<string, Expression<Action<RibbonFactory>>>
                                                          {
                                                              {"getVisible", f => f.GetVisible(null)},
                                                              {"getSupertip", f => f.GetSuperTip(null)},
                                                              {"getScreentip", f => f.GetScreenTip(null)},
                                                              {"getSize", f => f.GetSize(null)},
                                                              {"getKeytip", f => f.GetKeyTip(null)},
                                                              {"getLabel", f => f.GetLabel(null)},
                                                              {"getImageMso", f => f.GetImageMso(null)},
                                                              {"getImage", f => f.GetImage(null)},
                                                              {"getEnabled", f => f.GetEnabled(null)},
                                                              {"getDescription", f => f.GetDescription(null)}
                                                          }
                       },
                       {
                           "button", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                     {
                                         {"onAction", f => f.OnAction(null)},
                                         {"getShowLabel", f => f.GetShowLabel(null)},
                                         {"getShowImage", f => f.GetShowImage(null)}
                                     }
                       },
                       {
                           "checkBox", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                       {
                                           {"onAction", f => f.PressedOnAction(null, true)},
                                           {"getPressed", f => f.GetPressed(null)}
                                       }
                       },
                       {
                           "menu", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                   {
                                       {"onAction", f => f.OnAction(null)},
                                       {"getShowLabel", f => f.GetShowLabel(null)},
                                       {"getShowImage", f => f.GetShowImage(null)}
                                   }
                       },
                       {
                           "dropDown", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                       {
                                           {"getItemCount", f => f.GetItemCount(null)},
                                           {"getItemID", f => f.GetItemId(null, 0)},
                                           {"getItemImage", f => f.GetItemImage(null, 0)},
                                           {"getItemLabel", f => f.GetItemLabel(null, 0)},
                                           {"getItemScreenTip", f => f.GetItemScreenTip(null, 0)},
                                           {"getItemSuperTip", f => f.GetItemSuperTip(null, 0)},
                                           {"getSelectedItemID", f => f.GetSelectedItemId(null)},
                                           {"getSelectedItemIndex", f => f.GetSelectedItemIndex(null)},
                                           {"onAction", f => f.SelectionOnAction(null, null, 0)}
                                       }
                       },
                       {
                           "dynamicMenu", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                          {
                                              {"getContent", f => f.GetContent(null)}
                                          }
                       },
                       {
                           "gallery", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                      {
                                          {"getItemCount", f => f.GetItemCount(null)},
                                          {"getItemHeight", f => f.GetItemHeight(null)},
                                          {"getItemID", f => f.GetItemId(null, 0)},
                                          {"getItemImage", f => f.GetItemImage(null, 0)},
                                          {"getItemLabel", f => f.GetItemLabel(null, 0)},
                                          {"getItemScreenTip", f => f.GetItemScreenTip(null, 0)},
                                          {"getItemSuperTip", f => f.GetItemSuperTip(null, 0)},
                                          {"getSelectedItemID", f => f.GetSelectedItemId(null)},
                                          {"getSelectedItemIndex", f => f.GetSelectedItemIndex(null)},
                                          {"onAction", f => f.SelectionOnAction(null, null, 0)}
                                      }
                       },
                       {
                           "menuSeparator", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                            {
                                                {"getTitle", f => f.GetTitle(null)}
                                            }
                       },
                       {
                           "toggleButton", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                           {
                                               {"getPressed", f => f.GetPressed(null)},
                                               {"onAction", f => f.PressedOnAction(null, true)}
                                           }
                       },
                       {
                           "comboBox", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                       {
                                           {"getItemCount", f => f.GetItemCount(null)},
                                           {"getItemID", f => f.GetItemId(null, 0)},
                                           {"getItemImage", f => f.GetItemImage(null, 0)},
                                           {"getItemLabel", f => f.GetItemLabel(null, 0)},
                                           {"getItemScreenTip", f => f.GetItemScreenTip(null, 0)},
                                           {"getItemSuperTip", f => f.GetItemSuperTip(null, 0)},
                                           {"getText", f => f.GetText(null)},
                                           {"onChange", f => f.OnTextChanged(null, null)}
                                       }
                       },
                       {
                           "editBox", new Dictionary<string, Expression<Action<RibbonFactory>>>
                                      {
                                          {"getText", f => f.GetText(null)},
                                          {"onChange", f => f.OnTextChanged(null, null)}
                                      }
                       }
                   };
        }
    }
}