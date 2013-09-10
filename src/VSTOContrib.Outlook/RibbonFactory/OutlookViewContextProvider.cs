using Microsoft.Office.Interop.Outlook;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Outlook.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
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
        /// <typeparam name="TRibbonType">The type of the ribbon type.</typeparam>
        /// <param name="view">The view.</param>
        /// <returns></returns>
        public TRibbonType GetRibbonTypeForView<TRibbonType>(object view)
        {
            if (view is Explorer)
                return (TRibbonType)(object)OutlookRibbonType.OutlookExplorer;

            var selection = view as Selection;
            if (selection != null)
                return GetRibbonTypeForView<TRibbonType>(selection.Parent);

            return (TRibbonType) (object) InspectorToRibbonTypeConverter.Convert((Inspector) view);
        }
    }
}