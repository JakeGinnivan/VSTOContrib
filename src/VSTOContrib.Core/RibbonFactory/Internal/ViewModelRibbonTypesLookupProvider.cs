using System;
using System.Collections.Generic;
using System.Linq;
using VSTOContrib.Core.Annotations;

namespace VSTOContrib.Core.RibbonFactory.Internal
{
    class ViewModelRibbonTypesLookupProvider
    {
        readonly Dictionary<Type, string[]> viewModelRibbonTypes = new Dictionary<Type, string[]>();
        static ViewModelRibbonTypesLookupProvider instance;

        public static ViewModelRibbonTypesLookupProvider Instance
        {
            get { return instance ?? (instance = new ViewModelRibbonTypesLookupProvider()); }
        }

        public string[] GetRibbonTypesFor(Type ribbonViewModel, [CanBeNull] string fallbackType)
        {
            lock (viewModelRibbonTypes)
            {
                if (viewModelRibbonTypes.ContainsKey(ribbonViewModel))
                    return viewModelRibbonTypes[ribbonViewModel];

                var viewModelTypes = ribbonViewModel
                    .GetCustomAttributes(typeof(RibbonViewModelAttribute), false)
                    .OfType<RibbonViewModelAttribute>()
                    .Select(r => r.Type)
                    .ToArray();

                var ribbonTypesDefined = viewModelTypes.Any();
                if (!ribbonTypesDefined && fallbackType == null)
                    throw new InvalidOperationException("All IRibbonViewModel's must be marked up with a RibbonViewModel type. For example [OutlookRibbonViewModel(OutlookRibbonType.OutlookContact)]");

                var ribbonTypes = ribbonTypesDefined ? viewModelTypes : new[] {fallbackType};
                viewModelRibbonTypes.Add(ribbonViewModel, ribbonTypes);

                return viewModelRibbonTypes[ribbonViewModel];
            }
        }
    }
}
