using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using Microsoft.Office.Core;
using stdole;

namespace Outlook.Utility.RibbonFactory
{
    // I know partial classes are bad, but this contains all the callbacks. 
    // Moved into partial class to remove clutter from actual factory
    partial class RibbonFactory
    {
        /// <summary>
        /// button onAction callback
        /// </summary>
        /// <param name="control"></param>
        public void OnAction(IRibbonControl control)
        {
            Invoke(control, () => OnAction(null));
        }

        /// <summary>
        /// dropDown and gallery onAction callback
        /// </summary>
        /// <param name="control"></param>
        /// <param name="selectedId"></param>
        /// <param name="selectedIndex"></param>
        public void SelectionOnAction(IRibbonControl control, string selectedId, int selectedIndex)
        {
            Invoke(control, () => SelectionOnAction(null, null, 0), selectedId, selectedIndex);
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
            return (RibbonControlSize)Invoke(control, () => GetSize(null));
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
            return (IPictureDisp)Invoke(control, () => GetImage(null));
        }

        /// <summary>
        /// GetItemImage
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public IPictureDisp GetItemImage(IRibbonControl control, int index)
        {
            return (IPictureDisp)Invoke(control, () => GetItemImage(null, 0), index);
        }

        /// <summary>
        /// OnTextChanged callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="text">The text.</param>
        public void OnTextChanged(IRibbonControl control, string text)
        {
            Invoke(control, () => OnTextChanged(null, null), text);
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
