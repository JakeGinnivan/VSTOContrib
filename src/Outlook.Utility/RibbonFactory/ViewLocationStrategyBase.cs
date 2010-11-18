using System.IO;
using System.Reflection;

namespace Office.Utility
{
    ///<summary>
    /// The strategy to use to resolve the view
    ///</summary>
    public abstract class ViewLocationStrategyBase
    {
        ///<summary>
        /// Fetches the Ribbon XML for a given view
        ///</summary>
        ///<typeparam name="T">The View model to fetch the Ribbon XML for</typeparam>
        ///<returns>Ribbon XML</returns>
        public abstract string LocateViewForViewModel<T>() where T : RibbonViewModelBase;

        protected static string GetResourceText(string resourceName, Assembly viewAssembly)
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
    }
}