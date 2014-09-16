using Microsoft.Office.Interop.Outlook;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Outlook.RibbonFactory
{
    public class OutlookViewContextProvider : IViewContextProvider
    {
        readonly IOfficeApplicationEvents officeApplicationEvents;

        public OutlookViewContextProvider(IOfficeApplicationEvents officeApplicationEvents)
        {
            this.officeApplicationEvents = officeApplicationEvents;
        }

        /// <summary>
        /// Gets the context for view.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <returns></returns>
        public object GetContextForView(OfficeWin32Window view)
        {
            var inspector = view.Window as Inspector;
            if (inspector != null)
                return inspector.CurrentItem;

            var explorer = view.Window as Explorer;
            if (explorer != null)
                return explorer;

            var selection = view.Window as Selection;
            if (selection != null)
                return GetContextForView(officeApplicationEvents.ToOfficeWindow(selection.Parent));

            return null;
        }

        /// <summary>
        /// Gets the ribbon type for view.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <returns></returns>
        public string GetRibbonTypeForView(OfficeWin32Window view)
        {
            if (view.Window is Explorer)
                return OutlookRibbonType.OutlookExplorer.GetEnumDescription();

            var selection = view.Window as Selection;
            if (selection != null)
                return GetRibbonTypeForView(selection.Parent);

            return InspectorToRibbonTypeConverter.Convert((Inspector) view.Window).GetEnumDescription();
        }
    }
}