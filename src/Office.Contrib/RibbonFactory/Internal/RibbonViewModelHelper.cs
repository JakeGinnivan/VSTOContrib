using System;
using System.Collections.Generic;
using System.Linq;

namespace Office.Contrib.RibbonFactory.Internal
{
    internal class RibbonViewModelHelper
    {
        private readonly Dictionary<Type, IEnumerable<object>> _ribbonTypes
            = new Dictionary<Type, IEnumerable<object>>();

        public IEnumerable<TRibbonTypes> GetRibbonTypesFor<TRibbonTypes>(Type ribbonViewModel) where TRibbonTypes : struct 
        {
            var enumType = typeof(TRibbonTypes);

            if (!enumType.IsEnum) throw new ArgumentException("TRibbonTypes must be a enum type");

            var viewModelMetaAttributes = ribbonViewModel.GetCustomAttributes(typeof(RibbonViewModelAttribute), false);

            if (viewModelMetaAttributes.Length == 0)
                throw new InvalidOperationException("All IRibbonViewModel's must be marked up with a RibbonViewModel");

            var viewModelMetaData = (RibbonViewModelAttribute)viewModelMetaAttributes[0];

            if (!_ribbonTypes.ContainsKey(enumType))
                _ribbonTypes.Add(enumType, Enum.GetValues(enumType).Cast<object>());

            return _ribbonTypes[enumType]
                .Where(value => ((int)viewModelMetaData.Type & (int)value) == (int)value)
                .Cast<TRibbonTypes>();
        }
    }
}
