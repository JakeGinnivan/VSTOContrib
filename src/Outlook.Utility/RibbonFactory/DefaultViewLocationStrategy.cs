using System.Linq;
using Office.Utility;

namespace Outlook.Utility.RibbonFactory
{
    /// <summary>
    /// Locates the view, default method is an xml resource using the following methods:
    /// Given class is: MyAddin/Ribbons/ContactsRibbonViewModel.cs
    /// Will resolve in this order:
    /// MyAddin/Ribbons/ContactsRibbonViewModel.xml
    /// MyAddin/Ribbons/ContactsRibbonView.xml
    /// MyAddin/Ribbons/ContactsRibbon.xml
    /// ContactsRibbonViewModel.xml
    /// ContactsRibbonView.xml
    /// ContactsRibbon.xml
    /// </summary>
    /// <returns>Ribbon XML</returns>
    public class DefaultViewLocationStrategy : ViewLocationStrategyBase
    {
        public override string LocateViewForViewModel<T>()
        {
            var viewModelType = typeof (T);
            var viewAssembly = viewModelType.Assembly;

            var manifestResourceNames = viewAssembly.GetManifestResourceNames();
            var resources = manifestResourceNames.Where(r => r.EndsWith(".xml"));
            var viewName = viewModelType.Name;
            var exactName = viewName + ".xml";
            var noViewName = viewName.Replace("Model", string.Empty) + ".xml";
            var noViewModelName = viewName.Replace("ViewModel", string.Empty) + ".xml";
            var viewResource =
                resources.SingleOrDefault(r => r == viewModelType.Namespace + "." + exactName) ??
                resources.SingleOrDefault(r => r == viewModelType.Namespace + "." + noViewName) ??
                resources.SingleOrDefault(r => r == viewModelType.Namespace + "." + noViewModelName) ??
                resources.SingleOrDefault(r => r.EndsWith(exactName)) ??
                resources.SingleOrDefault(r => r.EndsWith(noViewName)) ??
                resources.SingleOrDefault(r => r.EndsWith(noViewModelName));
            if (viewResource == null)
                throw new ViewNotFoundException("Cannot locate view for " + viewModelType.FullName+ ". Make sure it is an Embedded Resource");

            return GetResourceText(viewResource, viewAssembly);
        }
    }
}