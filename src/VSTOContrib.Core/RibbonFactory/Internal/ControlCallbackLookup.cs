using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace VSTOContrib.Core.RibbonFactory.Internal
{
    internal class ControlCallbackLookup
    {
        /// <summary>
        /// Lookup: Control -> VSTO callback name -> Factory method name
        /// </summary>
        private readonly Dictionary<string, Dictionary<string, string>> controlCallbackLookup;
        /// <summary>
        /// Lookup: Control -> Factory method name -> VSTO callback name
        /// </summary>
        private readonly Dictionary<string, Dictionary<string, string>> controlReverseCallbackLookup;

        public ControlCallbackLookup()
        {
            controlCallbackLookup = new Dictionary<string, Dictionary<string, string>>();
            controlReverseCallbackLookup = new Dictionary<string, Dictionary<string, string>>();

            foreach (var control in controlCallbackLookupExpressions.Where(c => c.Key != RibbonFactory.CommonCallbacks))
            {
                controlCallbackLookup.Add(control.Key, new Dictionary<string, string>());
                controlReverseCallbackLookup.Add(control.Key, new Dictionary<string, string>());
                foreach (var controlCallbacks in control.Value.Concat(controlCallbackLookupExpressions[RibbonFactory.CommonCallbacks]))
                {
                    var methodName = controlCallbacks.Value.GetMethodName();
                    controlCallbackLookup[control.Key].Add(controlCallbacks.Key, methodName);
                    controlReverseCallbackLookup[control.Key].Add(methodName, controlCallbacks.Key);
                }
            }
        }

        public IEnumerable<string> RibbonControls
        {
            get { return controlCallbackLookup.Keys; }
        }

        public IEnumerable<string> GetVstoControlCallbacks(string control)
        {
            //Control specific callbacks + common callbacks
            return controlCallbackLookup[control].Keys;
        }

        public string GetFactoryMethodName(string control, string vstoCallback)
        {
            return controlCallbackLookup[control][vstoCallback];
        }

        public string GetVstoCallback(string control, string factoryMethod)
        {
            return controlReverseCallbackLookup[control][factoryMethod];
        }


        readonly IDictionary<string, Dictionary<string, Expression<Action<RibbonFactory>>>>
            controlCallbackLookupExpressions =
                new Dictionary<string, Dictionary<string, Expression<Action<RibbonFactory>>>>
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
                        "group", new Dictionary<string, Expression<Action<RibbonFactory>>>
                        {
                            {"getHelperText", f => f.GetHelperText(null)}
                        }
                    },
                    {"tab", new Dictionary<string, Expression<Action<RibbonFactory>>>()},
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