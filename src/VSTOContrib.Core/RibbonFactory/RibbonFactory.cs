using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using VSTOContrib.Core.Annotations;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;

namespace VSTOContrib.Core.RibbonFactory
{
    /// <summary>
    /// Simplifies adding custom Ribbon's to Office. 
    /// Allows the custom Ribbon xml to be wired up to IRibbonViewModel's
    /// by convention. Simply name the Ribbon.xml the same as the ribbon view model class
    /// in the same assembly
    /// </summary>
    [ComVisible(true)]
    public abstract class RibbonFactory : IRibbonFactory
    {
        internal const string CommonCallbacks = "CommonCallbacks";
        readonly IRibbonFactoryController ribbonFactoryController;
        readonly VstoContribContext context;

        protected RibbonFactory(AddInBase addinBase, Assembly[] assemblies, IViewContextProvider contextProvider, IOfficeApplicationEvents officeApplicationEvents, [CanBeNull] string fallbackRibbonType)
        {
            if (assemblies.Length == 0)
                throw new InvalidOperationException("You must specify at least one assembly to scan for viewmodels");

            context = new VstoContribContext(assemblies, addinBase, fallbackRibbonType);
            ribbonFactoryController = new RibbonFactoryController(contextProvider, context, officeApplicationEvents);

            addinBase.Startup += AddinBaseOnStartup;
            addinBase.Shutdown += AddinBaseOnShutdown;
        }

        protected static Assembly[] UseIfEmpty(Assembly[] assemblies, Assembly fallback)
        {
            if (assemblies.Length > 0)
                return assemblies;

            return new[] { fallback };
        }

        void AddinBaseOnStartup(object sender, EventArgs eventArgs)
        {
            VstoContribLog.Debug(_ => _("AddinBase.Startup raised, initialising ribbon factory controller"));
            context.AddinBase.Startup -= AddinBaseOnStartup;

            InitialiseRibbonFactoryController(ribbonFactoryController, context.Application);
        }

        void AddinBaseOnShutdown(object sender, EventArgs eventArgs)
        {
            context.AddinBase.Shutdown -= AddinBaseOnShutdown;
            ribbonFactoryController.Dispose();
            ShuttingDown();
        }

        ///<summary>
        /// Gets or Sets the strategy that fetches the Ribbon XML for a given view
        ///</summary>
        public IViewLocationStrategy LocateViewStrategy
        {
            get { return context.ViewLocationStrategy; }
        }

        public IViewModelFactory ViewModelFactory
        {
            get { return context.ViewModelFactory; }
            set
            {
                if (value == null) return;
                context.ViewModelFactory = value;
            }
        }

        /// <summary>
        /// Gets collection of registered global error handlers. New custom IErrorHandler implementation can be added in collection.
        /// </summary>
        public ICollection<IErrorHandler> ErrorHandlers
        {
            get { return context.ErrorHandlers; }
        }

        /// <summary>
        /// Called when the add-in is shutting down
        /// </summary>
        protected abstract void ShuttingDown();

        /// <summary>
        /// Initialisation callback for ribbon factory. The implementation must initialise the controller and 
        /// </summary>
        protected abstract void InitialiseRibbonFactoryController(IRibbonFactoryController controller, object application);

        /// <summary>
        /// Ribbon_s the load.
        /// </summary>
        /// <param name="ribbonUi">The ribbon UI.</param>
        // ReSharper disable InconsistentNaming
        public virtual void Ribbon_Load(IRibbonUI ribbonUi)
        {
            VstoContribLog.Debug(_ => _("Ribbon_Load event raised: {0}", ribbonUi.ToLogFormat()));
            ribbonFactoryController.RibbonLoaded(ribbonUi);
        }
        // ReSharper restore InconsistentNaming

        /// <summary>
        /// Gets the custom UI.
        /// </summary>
        /// <param name="ribbonId">The ribbon id.</param>
        /// <returns></returns>
        public string GetCustomUI(string ribbonId)
        {
            VstoContribLog.Info(_ => _("Office called GetCustomUI({0})", ribbonId));
            var customUI = ribbonFactoryController.GetCustomUI(ribbonId);
            VstoContribLog.Debug(_ => _("Provided ribbon xml for ribbonId {0}:\r\n\r\n{1}", ribbonId, customUI));
            return customUI;
        }


        /***************************************************/
        /*                                                 */
        /*                   Callbacks                     */
        /*                                                 */
        /***************************************************/

        /// <summary>
        /// button onAction callback
        /// </summary>
        /// <param name="control"></param>
        public void OnAction(IRibbonControl control)
        {
            ribbonFactoryController.Invoke(control, () => OnAction(null));
        }

        /// <summary>
        /// dropDown and gallery onAction callback
        /// </summary>
        /// <param name="control"></param>
        /// <param name="selectedId"></param>
        /// <param name="selectedIndex"></param>
        public void SelectionOnAction(IRibbonControl control, string selectedId, int selectedIndex)
        {
            ribbonFactoryController.Invoke(control, () => SelectionOnAction(null, null, 0), selectedId, selectedIndex);
        }

        /// <summary>
        /// checkBox and togglebutton onAction callback
        /// </summary>
        /// <param name="control"></param>
        /// <param name="pressed"></param>
        public void PressedOnAction(IRibbonControl control, bool pressed)
        {
            ribbonFactoryController.Invoke(control, () => PressedOnAction(null, true), pressed);
        }

        /// <summary>
        /// GetDescription callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetDescription(IRibbonControl control)
        {
            return (string)ribbonFactoryController.InvokeGet(control, () => GetDescription(null));
        }

        /// <summary>
        /// GetEnabled callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public bool GetEnabled(IRibbonControl control)
        {
            return (bool)ribbonFactoryController.InvokeGet(control, () => GetEnabled(null));
        }

        /// <summary>
        /// GetImageMso callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetImageMso(IRibbonControl control)
        {
            return (string)ribbonFactoryController.InvokeGet(control, () => GetImageMso(null));
        }

        /// <summary>
        /// GetLabel callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetLabel(IRibbonControl control)
        {
            return (string)ribbonFactoryController.InvokeGet(control, () => GetLabel(null));
        }

        /// <summary>
        /// GetKeyTip callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetKeyTip(IRibbonControl control)
        {
            return (string)ribbonFactoryController.InvokeGet(control, () => GetKeyTip(null));
        }

        /// <summary>
        /// GetScreenTip
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetScreenTip(IRibbonControl control)
        {
            return (string)ribbonFactoryController.InvokeGet(control, () => GetScreenTip(null));
        }

        /// <summary>
        /// GetSuperTip
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetSuperTip(IRibbonControl control)
        {
            return (string)ribbonFactoryController.InvokeGet(control, () => GetSuperTip(null));
        }

        /// <summary>
        /// GetVisible callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public bool GetVisible(IRibbonControl control)
        {
            return (bool)ribbonFactoryController.InvokeGet(control, () => GetVisible(null));
        }

        /// <summary>
        /// GetShowImage callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public bool GetShowImage(IRibbonControl control)
        {
            return (bool)ribbonFactoryController.InvokeGet(control, () => GetShowImage(null));
        }

        /// <summary>
        /// GetShowLabel
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public bool GetShowLabel(IRibbonControl control)
        {
            return (bool)ribbonFactoryController.InvokeGet(control, () => GetShowLabel(null));
        }

        /// <summary>
        /// GetItemCount callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public int GetItemCount(IRibbonControl control)
        {
            return (int)ribbonFactoryController.InvokeGet(control, () => GetItemCount(null));
        }

        /// <summary>
        /// GetItemId callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public string GetItemId(IRibbonControl control, int index)
        {
            return (string)ribbonFactoryController.InvokeGet(control, () => GetItemId(null, index), index);
        }

        /// <summary>
        /// GetItemLabel callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public string GetItemLabel(IRibbonControl control, int index)
        {
            return (string)ribbonFactoryController.InvokeGet(control, () => GetItemLabel(null, index), index);
        }

        /// <summary>
        /// GetItemScreenTip callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public string GetItemScreenTip(IRibbonControl control, int index)
        {
            return (string)ribbonFactoryController.InvokeGet(control, () => GetItemScreenTip(null, index), index);
        }

        /// <summary>
        /// GetItemSuperTip callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public string GetItemSuperTip(IRibbonControl control, int index)
        {
            return (string)ribbonFactoryController.InvokeGet(control, () => GetItemSuperTip(null, index), index);
        }

        /// <summary>
        /// GetSelectedItemId callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetSelectedItemId(IRibbonControl control)
        {
            return (string)ribbonFactoryController.InvokeGet(control, () => GetSelectedItemId(null));
        }

        /// <summary>
        /// GetSelectedItemIndex callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public int GetSelectedItemIndex(IRibbonControl control)
        {
            return (int)ribbonFactoryController.InvokeGet(control, () => GetSelectedItemIndex(null));
        }

        /// <summary>
        /// GetContent callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetContent(IRibbonControl control)
        {
            return (string)ribbonFactoryController.InvokeGetContent(control, () => GetContent(null));
        }

        /// <summary>
        /// GetText callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetText(IRibbonControl control)
        {
            return (string)ribbonFactoryController.InvokeGet(control, () => GetText(null));
        }

        /// <summary>
        /// GetTitle callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public string GetTitle(IRibbonControl control)
        {
            return (string)ribbonFactoryController.InvokeGet(control, () => GetTitle(null));
        }

        /// <summary>
        /// GetPressed callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public bool GetPressed(IRibbonControl control)
        {
            return (bool)ribbonFactoryController.InvokeGet(control, () => GetPressed(null));
        }

        /// <summary>
        /// GetSize callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public RibbonControlSize GetSize(IRibbonControl control)
        {
            return (RibbonControlSize)ribbonFactoryController.InvokeGet(control, () => GetSize(null));
        }

        /// <summary>
        /// GetItemHeight
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public int GetItemHeight(IRibbonControl control)
        {
            return (int)ribbonFactoryController.InvokeGet(control, () => GetItemHeight(control));
        }

#if OFFICE2007
        /// <summary>
        /// GetImage
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public stdole.IPictureDisp GetImage(IRibbonControl control)
        {
            return (stdole.IPictureDisp)ribbonFactoryController.InvokeGet(control, () => GetImage(null));
        }

        /// <summary>
        /// GetItemImage
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public stdole.IPictureDisp GetItemImage(IRibbonControl control, int index)
        {
            return (stdole.IPictureDisp)ribbonFactoryController.InvokeGet(control, () => GetItemImage(null, 0), index);
        }
#else
        /// <summary>
        /// GetImage
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns></returns>
        public Bitmap GetImage(IRibbonControl control)
        {
            return (Bitmap)ribbonFactoryController.InvokeGet(control, () => GetImage(null));
        }

        /// <summary>
        /// GetItemImage
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        public Bitmap GetItemImage(IRibbonControl control, int index)
        {
            return (Bitmap)ribbonFactoryController.InvokeGet(control, () => GetItemImage(null, 0), index);
        }
#endif
        /// <summary>
        /// OnTextChanged callback
        /// </summary>
        /// <param name="control">The control.</param>
        /// <param name="text">The text.</param>
        public void OnTextChanged(IRibbonControl control, string text)
        {
            ribbonFactoryController.Invoke(control, () => OnTextChanged(null, null), text);
        }

        /// <summary>
        /// GetHelperText callback
        /// </summary>
        /// <param name="control">The control</param>
        /// <returns></returns>
        public string GetHelperText(IRibbonControl control)
        {
            return (string)ribbonFactoryController.InvokeGet(control, () => GetHelperText(null));
        }
    }
}
