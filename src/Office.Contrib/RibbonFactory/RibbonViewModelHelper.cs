using System;
using System.Collections.Generic;
using System.Linq;

namespace Office.Contrib.RibbonFactory
{
    internal static class RibbonViewModelHelper
    {
        private static readonly Dictionary<Type, IEnumerable<object>> RibbonTypes
            = new Dictionary<Type, IEnumerable<object>>();

        public static IEnumerable<TRibbonTypes> GetRibbonTypesFor<TRibbonTypes>(Type ribbonViewModel) where TRibbonTypes : struct 
        {
            var enumType = typeof(TRibbonTypes);

            if (!enumType.IsEnum) throw new ArgumentException("TRibbonTypes must be a enum type");

            var viewModelMetaAttributes = ribbonViewModel.GetCustomAttributes(typeof(RibbonViewModelAttribute), false);

            if (viewModelMetaAttributes.Length == 0)
                throw new InvalidOperationException("All IRibbonViewModel's must be marked up with a RibbonViewModel");

            var viewModelMetaData = (RibbonViewModelAttribute)viewModelMetaAttributes[0];

            if (!RibbonTypes.ContainsKey(enumType))
                RibbonTypes.Add(enumType, Enum.GetValues(enumType).Cast<object>());

            return RibbonTypes.Cast<TRibbonTypes>().Where(value =>
                    ((int)viewModelMetaData.Type & (int)(object)value) == (int)(object)value);
        }
    }
}
