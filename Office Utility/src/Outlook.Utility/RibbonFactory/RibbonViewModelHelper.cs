using System;
using System.Collections.Generic;
using System.Linq;

namespace Outlook.Utility.RibbonFactory
{
    internal static class RibbonViewModelHelper
    {
        private static readonly IEnumerable<RibbonType> RibbonTypes;

        static RibbonViewModelHelper()
        {
            RibbonTypes = Enum.GetValues(typeof(RibbonType)).Cast<RibbonType>();            
        }

        public static IEnumerable<RibbonType> GetRibbonTypesFor(Type ribbonViewModel)
        {
            var viewModelMetaAttributes = ribbonViewModel.GetCustomAttributes(typeof(RibbonViewModelAttribute), false);

            if (viewModelMetaAttributes.Length == 0)
                throw new InvalidOperationException("All IRibbonViewModel's must be marked up with a RibbonViewModel");

            var viewModelMetaData = (RibbonViewModelAttribute)viewModelMetaAttributes[0];

            return RibbonTypes.Where(value => (viewModelMetaData.Type & value) == value);
        }
    }
}
