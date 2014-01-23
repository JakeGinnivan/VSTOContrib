using Microsoft.Office.Interop.Outlook;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Outlook.RibbonFactory
{
    public class OutlookViewContextProvider : IViewContextProvider
    {
        /// <summary>
        /// Gets the context for view.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <returns></returns>
        public object GetContextForView(object view)
        {
            var inspector = view as Inspector;
            if (inspector != null)
                return inspector.CurrentItem;

            var explorer = view as Explorer;
            if (explorer != null)
                return explorer;

            var selection = view as Selection;
            if (selection != null)
                return GetContextForView(selection.Parent);

            return null;
        }

        /// <summary>
        /// Gets the ribbon type for view.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <returns></returns>
        public string GetRibbonTypeForView(object view)
        {
            if (view is Explorer)
                return OutlookRibbonType.OutlookExplorer.GetEnumDescription();

            var selection = view as Selection;
            if (selection != null)
                return GetRibbonTypeForView(selection.Parent);

            return InspectorToRibbonTypeConverter.Convert((Inspector) view).GetEnumDescription();
        }
    }
}