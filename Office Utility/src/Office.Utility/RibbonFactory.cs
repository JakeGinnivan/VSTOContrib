using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Core;
using stdole;

namespace Office.Utility
{
    /// <summary>
    /// Simplifies adding custom Ribbon's to Office. 
    /// Allows the custom Ribbon xml to be wired up to IRibbonViewModel's
    /// by convention. Simply name the Ribbon.xml the same as the ribbon view model class
    /// in the same assembly
    /// </summary>
    [ComVisible(true)]
    public class RibbonFactory : IRibbonFactory
    {
        private readonly Dictionary<RibbonType, IRibbonViewModel> _ribbons = new Dictionary<RibbonType, IRibbonViewModel>();
        private readonly Dictionary<Type, string> _ribbonViews = new Dictionary<Type, string>();
        private readonly Dictionary<string, CallbackTarget> _ribbonCallbackTarget = new Dictionary<string, CallbackTarget>();
        const string OfficeCustomui = "http://schemas.microsoft.com/office/2006/01/customui";
        const string OfficeCustomui4 = "http://schemas.microsoft.com/office/2009/07/customui";
        internal const string CommonCallbacks = "CommonCallbacks";
        private IRibbonViewModel _currentlyLoadingRibbon;
        private ControlCallbackLookup _controlCallbackLookup;
        private static RibbonFactory _instance;

        private RibbonFactory()
        { }


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
        /// <param name="ribbons">Ribbon view models to wire up</param>
        /// <exception cref="ViewNotFoundException">If the view cannot be located for a view model</exception>
        public void InitialiseFactory(IEnumerable<IRibbonViewModel> ribbons)
        {
            _controlCallbackLookup = new ControlCallbackLookup(GetRibbonElements());

            Expression<Action> loadMethod = () => Ribbon_Load(null);
            var loadMethodName = loadMethod.GetMethodName();

            foreach (var ribbonViewModel in ribbons)
            {
                _ribbons.Add(ribbonViewModel.Type, ribbonViewModel);
            }

            foreach (var ribbonViewModel in ribbons)
            {
                var viewModelType = ribbonViewModel.GetType();
                var resourceText = LocateView(viewModelType);

                var ribbonDoc = XDocument.Parse(resourceText);

                //We have to override the Ribbon_Load event to make sure we get the callback
                var customUi = 
                    ribbonDoc.Descendants(XName.Get("customUI", OfficeCustomui)).SingleOrDefault()
                    ?? ribbonDoc.Descendants(XName.Get("customUI", OfficeCustomui4)).Single();

                customUi.SetAttributeValue("onLoad", loadMethodName);

                WireUpEvents(viewModelType, ribbonDoc, ribbonViewModel);
                _ribbonViews.Add(viewModelType, ribbonDoc.ToString());
            }
        }

        /// <summary>
        /// Locates the view, default method is an xml resource with the same name and in the same namespace as the view.
        /// for example:
        /// MyAddin/Ribbons/ContactsRibbon.cs
        /// will look for
        /// MyAddin/Ribbons/ContactsRibbon.xml
        /// </summary>
        /// <param name="viewModelType">Type of the view model.</param>
        /// <returns>Ribbon XML</returns>
        protected virtual string LocateView(Type viewModelType)
        {
            var viewAssembly = viewModelType.Assembly;

            var resources = viewAssembly.GetManifestResourceNames();
            var viewName = viewModelType.Name;
            var viewResource =
                resources.SingleOrDefault(r => r == viewModelType.Namespace + "." + viewName + ".xml");
            if (viewResource == null)
                throw new ViewNotFoundException("Cannot locate view for " + viewModelType.FullName);

            return GetResourceText(viewResource, viewAssembly);
        }

        // ReSharper disable InconsistentNaming
        /// <summary>
        /// Ribbon_s the load.
        /// </summary>
        /// <param name="ribbonUi">The ribbon UI.</param>
        public void Ribbon_Load(IRibbonUI ribbonUi)
        {
            _currentlyLoadingRibbon.RibbonUi = ribbonUi;
        }
        // ReSharper restore InconsistentNaming

        private void WireUpEvents(Type viewModelType, XContainer ribbonDoc, IRibbonViewModel model)
        {
            //Go through each type of Ribbon 
            foreach (var ribbonControl in _controlCallbackLookup.RibbonControls)
            {
                //Get each instance of that control in the ribbon definition file
                var xElements = ribbonDoc.Descendants(XName.Get(ribbonControl, OfficeCustomui));

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
                        var callbackTag = BuildTag(viewModelType, elementId, factoryMethodName);
                        _ribbonCallbackTarget.Add(callbackTag, new CallbackTarget(model, currentCallback));
                        xElement.SetAttributeValue(XName.Get("tag"), (viewModelType.FullName + elementId.Value));
                    }
                }
            }
        }

        private static string BuildTag(Type viewModelType, XAttribute elementId, string factoryMethodName)
        {
            return viewModelType.FullName + elementId.Value + factoryMethodName;
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
            if (!_ribbons.ContainsKey(enumFromDescription)) return null;

            var ribbonViewModel = _ribbons[enumFromDescription];
            var type = ribbonViewModel.GetType();
            if (!_ribbonViews.ContainsKey(type)) return null;

            _currentlyLoadingRibbon = ribbonViewModel;
            return _ribbonViews[type];
        }

        /// <summary>
        /// button onAction callback
        /// </summary>
        /// <param name="control"></param>
        public void OnAction(IRibbonControl control)
        {
            Invoke(control, ()=>OnAction(null));
        }

        /// <summary>
        /// dropDown and gallery onAction callback
        /// </summary>
        /// <param name="control"></param>
        /// <param name="selectedId"></param>
        /// <param name="selectedIndex"></param>
        public void SelectionOnAction(IRibbonControl control, string selectedId, int selectedIndex)
        {
            Invoke(control, ()=>SelectionOnAction(null, null, 0), selectedId, selectedIndex);
        }

        /// <summary>
        /// checkBox and togglebutton onAction callback
        /// </summary>
        /// <param name="control"></param>
        /// <param name="pressed"></param>
        public void PressedOnAction(IRibbonControl control, bool pressed)
        {
            Invoke(control, () => PressedOnAction(null, true), pressed);
        }

        /// <summary>
        /// GetDescription callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetDescription(IRibbonControl control)
        {
            return (string)Invoke(control, () => GetDescription(null));
        }

        /// <summary>
        /// GetEnabled callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public bool GetEnabled(IRibbonControl control)
        {
            return (bool)Invoke(control, () => GetEnabled(null));
        }

        /// <summary>
        /// GetImageMso callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetImageMso(IRibbonControl control)
        {
            return (string)Invoke(control, () => GetImageMso(null));
        }

        /// <summary>
        /// GetLabel callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetLabel(IRibbonControl control)
        {
            return (string)Invoke(control, () => GetLabel(null));
        }

        /// <summary>
        /// GetKeyTip callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetKeyTip(IRibbonControl control)
        {
            return (string)Invoke(control, () => GetKeyTip(null));
        }

        /// <summary>
        /// GetScreenTip
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetScreenTip(IRibbonControl control)
        {
            return (string)Invoke(control, () => GetScreenTip(null));
        }

        /// <summary>
        /// GetSuperTip
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetSuperTip(IRibbonControl control)
        {
            return (string)Invoke(control, () => GetSuperTip(null));
        }

        /// <summary>
        /// GetVisible callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public bool GetVisible(IRibbonControl control)
        {
            return (bool)Invoke(control, () => GetVisible(null));
        }

        /// <summary>
        /// GetShowImage callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public bool GetShowImage(IRibbonControl control)
        {
            return (bool)Invoke(control, () => GetShowImage(null));
        }

        /// <summary>
        /// GetShowLabel
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public bool GetShowLabel(IRibbonControl control)
        {
            return (bool)Invoke(control, () => GetShowLabel(null));
        }

        /// <summary>
        /// GetItemCount callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public int GetItemCount(IRibbonControl control)
        {
            return (int)Invoke(control, () => GetItemCount(null));
        }

        /// <summary>
        /// GetItemId callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public string GetItemId(IRibbonControl control, int index)
        {
            return (string)Invoke(control, () => GetItemId(null, 0));
        }

        /// <summary>
        /// GetItemLabel callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public string GetItemLabel(IRibbonControl control, int index)
        {
            return (string)Invoke(control, () => GetItemLabel(null, 0));
        }

        /// <summary>
        /// GetItemScreenTip callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public string GetItemScreenTip(IRibbonControl control, int index)
        {
            return (string)Invoke(control, () => GetItemScreenTip(null, 0));
        }

        /// <summary>
        /// GetItemSuperTip callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public string GetItemSuperTip(IRibbonControl control, int index)
        {
            return (string)Invoke(control, () => GetItemSuperTip(null, 0));
        }

        /// <summary>
        /// GetSelectedItemId callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public int GetSelectedItemId(IRibbonControl control)
        {
            return (int)Invoke(control, () => GetSelectedItemId(null));
        }

        /// <summary>
        /// GetSelectedItemIndex callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public int GetSelectedItemIndex(IRibbonControl control)
        {
            return (int)Invoke(control, () => GetSelectedItemIndex(null));
        }

        /// <summary>
        /// GetContent callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetContent(IRibbonControl control)
        {
            return (string)Invoke(control, () => GetContent(null));
        }

        /// <summary>
        /// GetText callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetText(IRibbonControl control)
        {
            return (string)Invoke(control, () => GetText(null));
        }

        /// <summary>
        /// GetTitle callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetTitle(IRibbonControl control)
        {
            return (string)Invoke(control, () => GetTitle(null));
        }

        /// <summary>
        /// GetPressed callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public bool GetPressed(IRibbonControl control)
        {
            return (bool)Invoke(control, () => GetPressed(null));
        }

        /// <summary>
        /// GetSize callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public RibbonControlSize GetSize(IRibbonControl control)
        {
            return (RibbonControlSize)Invoke(control, ()=>GetSize(null));
        }

        /// <summary>
        /// GetItemHeight
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public int GetItemHeight(IRibbonControl control)
        {
            return (int)Invoke(control, () => GetItemHeight(control));
        }

        /// <summary>
        /// GetImage
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public IPictureDisp GetImage(IRibbonControl control)
        {
            return (IPictureDisp)Invoke(control, ()=>GetImage(null));
        }

        /// <summary>
        /// GetItemImage
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public IPictureDisp GetItemImage(IRibbonControl control, int index)
        {
            return (IPictureDisp)Invoke(control, ()=>GetItemImage(null, 0), index);
        }

        /// <summary>
        /// OnTextChanged callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="text">The text.</param>
        public void OnTextChanged(IRibbonControl control, string text)
        {
            Invoke(control, ()=>OnTextChanged(null, null), text);
        }

        private object Invoke(IRibbonControl control, Expression<Action> caller, params object[] parameters)
        {
            var callbackTarget = _ribbonCallbackTarget[control.Tag+caller.GetMethodName()];
            var viewModelType = callbackTarget.ViewModel.GetType();

            return viewModelType.InvokeMember(callbackTarget.Method,
                                       BindingFlags.InvokeMethod,
                                       null,
                                       callbackTarget.ViewModel,
                                       new[]
                                           {
                                               control
                                           }
                                       .Concat(parameters)
                                       .ToArray());
        }

        private static string GetResourceText(string resourceName, Assembly viewAssembly)
        {
            using (var stream = viewAssembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return null;
                using (var resourceReader = new StreamReader(stream))
                {
                    return resourceReader.ReadToEnd();
                }
            }
        }

        private Dictionary<string, Dictionary<string, Expression<Action>>> GetRibbonElements()
        {
            return new Dictionary<string, Dictionary<string, Expression<Action>>>
                                  {
                                      {CommonCallbacks, new Dictionary<string, Expression<Action>>
                                                            {
                                                                {"getVisible", ()=>GetVisible(null)},
                                                                {"getSupertip", ()=>GetSuperTip(null)},
                                                                {"getScreentip", ()=>GetScreenTip(null)},
                                                                {"getSize", ()=>GetSize(null)},
                                                                {"getKeytip", ()=>GetKeyTip(null)},
                                                                {"getLabel", ()=>GetLabel(null)},
                                                                {"getImageMso",()=>GetImageMso(null)},
                                                                {"getImage", ()=>GetImage(null)},
                                                                {"getEnabled", ()=>GetEnabled(null)},
                                                                {"getDescription", ()=>GetDescription(null)}
                                                            }},
                                      {"button", new Dictionary<string, Expression<Action>>
                                                     {
                                                         {"onAction", ()=>OnAction(null)},
                                                         {"getShowLabel", ()=>GetShowLabel(null)},
                                                         {"getShowImage", ()=>GetShowImage(null)}
                                                     }},
                                      {"checkBox", new Dictionary<string, Expression<Action>>
                                                       {
                                                           {"onAction", ()=>PressedOnAction(null, true)},
                                                           {"getPressed", ()=>GetPressed(null)}
                                                       }},
                                      {"dropDown", new Dictionary<string, Expression<Action>>
                                                       {
                                                           {"getItemCount", ()=>GetItemCount(null)},
                                                           {"getItemID", ()=>GetItemId(null, 0)},
                                                           {"getItemImage", ()=>GetItemImage(null, 0)},
                                                           {"getItemLabel", ()=>GetItemLabel(null, 0)},
                                                           {"getItemScreenTip", ()=>GetItemScreenTip(null, 0)},
                                                           {"getItemSuperTip", ()=>GetItemSuperTip(null, 0)},
                                                           {"getSelectedItemID", ()=>GetSelectedItemId(null)},
                                                           {"getSelectedItemIndex", ()=>GetSelectedItemIndex(null)},
                                                           {"onAction", ()=>SelectionOnAction(null, null, 0)}
                                                       }},
                                      {"dynamicMenu", new Dictionary<string, Expression<Action>>
                                                          {
                                                              {"getContent", ()=>GetContent(null)}
                                                          }},
                                      {"gallery", new Dictionary<string, Expression<Action>>
                                                      {
                                                          {"getItemCount", ()=>GetItemCount(null)},
                                                          {"getItemHeight", ()=>GetItemHeight(null)},
                                                          {"getItemID", ()=>GetItemId(null, 0)},
                                                          {"getItemImage", ()=>GetItemImage(null, 0)},
                                                          {"getItemLabel", ()=>GetItemLabel(null, 0)},
                                                          {"getItemScreenTip", ()=>GetItemScreenTip(null, 0)},
                                                          {"getItemSuperTip", ()=>GetItemSuperTip(null, 0)},
                                                          {"getSelectedItemID", ()=>GetSelectedItemId(null)},
                                                          {"getSelectedItemIndex", ()=>GetSelectedItemIndex(null)},
                                                          {"onAction", ()=>SelectionOnAction(null, null, 0)}
                                                      }},
                                      {"menuSeparator",  new Dictionary<string, Expression<Action>>
                                                             {
                                                                 {"getTitle", ()=>GetTitle(null)}
                                                             }},
                                      {"toggleButton", new Dictionary<string, Expression<Action>>
                                                           {
                                                               {"getPressed", ()=>GetPressed(null)},
                                                               {"onAction", ()=>PressedOnAction(null, true)}
                                                           }},
                                      {"comboBox", new Dictionary<string, Expression<Action>>
                                                       {
                                                           {"getItemCount", ()=>GetItemCount(null)},
                                                           {"getItemID", ()=>GetItemId(null, 0)},
                                                           {"getItemImage", ()=>GetItemImage(null, 0)},
                                                           {"getItemLabel", ()=>GetItemLabel(null, 0)},
                                                           {"getItemScreenTip", ()=>GetItemScreenTip(null, 0)},
                                                           {"getItemSuperTip", ()=>GetItemSuperTip(null, 0)},
                                                           {"getText", ()=>GetText(null)},
                                                           {"onChange", ()=>OnTextChanged(null, null)}
                                                       }},
                                      {"editBox", new Dictionary<string, Expression<Action>>
                                                      {
                                                          {"getText", ()=>GetText(null)},
                                                          {"onChange", ()=>OnTextChanged(null, null)}
                                                      }}
                                  };
        }

        
    }
}
