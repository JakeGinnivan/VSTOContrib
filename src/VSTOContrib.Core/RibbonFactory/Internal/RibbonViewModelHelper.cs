using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace VSTOContrib.Core.RibbonFactory.Internal
{
    internal class RibbonViewModelHelper
    {
        private readonly Dictionary<Type, object> viewModelRibbonTypes = new Dictionary<Type, object>();

        public IEnumerable<TRibbonTypes> GetRibbonTypesFor<TRibbonTypes>(Type ribbonViewModel) where TRibbonTypes : struct 
        {
            var enumType = typeof(TRibbonTypes);

            if (!enumType.IsEnum) throw new ArgumentException("TRibbonTypes must be a enum type");

            if (viewModelRibbonTypes.ContainsKey(ribbonViewModel)) return (TRibbonTypes[])viewModelRibbonTypes[ribbonViewModel];

            var viewModelMetaAttributes = (RibbonViewModelAttribute)ribbonViewModel.GetCustomAttributes(typeof(RibbonViewModelAttribute), false).SingleOrDefault();
            var viewModelType = GetRibbonTypeAttributeValue<TRibbonTypes>(enumType, viewModelMetaAttributes);

            if (viewModelType == null)
                throw new InvalidOperationException("All IRibbonViewModel's must be marked up with a RibbonViewModel type. For example [OutlookRibbonViewModel(OutlookRibbonType.OutlookContact)]");

            if (!viewModelRibbonTypes.ContainsKey(ribbonViewModel))
            {
                var ribbonTypesFor = Enum.GetValues(enumType).Cast<object>()
                    .Where(value => (viewModelType & (int)value) == (int)value)
                    .Cast<TRibbonTypes>()
                    .ToArray();

                viewModelRibbonTypes.Add(ribbonViewModel, ribbonTypesFor);
            }

            return (TRibbonTypes[]) viewModelRibbonTypes[ribbonViewModel];
        }

        static int? GetRibbonTypeAttributeValue<TRibbonTypes>(Type enumType, RibbonViewModelAttribute viewModelMetaAttributes) where TRibbonTypes : struct
        {
            var defaultValue = new Lazy<TRibbonTypes?>(() =>
            {
                var defaultAttribute = (DefaultValueAttribute)enumType.GetCustomAttributes(typeof (DefaultValueAttribute), false).SingleOrDefault();
                if (defaultAttribute != null)
                    return (TRibbonTypes) defaultAttribute.Value;
                return null;
            });
            int? viewModelType;

            if (viewModelMetaAttributes == null && !defaultValue.Value.HasValue)
                viewModelType = null;
            else if (viewModelMetaAttributes != null)
                viewModelType = (int)viewModelMetaAttributes.Type;
            else
                viewModelType = (int) (object) defaultValue.Value.Value;

            return viewModelType;
        }
    }
}
