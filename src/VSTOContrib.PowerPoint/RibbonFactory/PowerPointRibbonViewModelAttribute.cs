using VSTOContrib.Core.RibbonFactory;

namespace VSTOContrib.PowerPoint.RibbonFactory
{
    public class PowerPointRibbonViewModelAttribute : RibbonViewModelAttribute
    {
        public PowerPointRibbonViewModelAttribute()
            : base(PowerPointRibbonType.PowerPointPresentation)
        {
        }

        public PowerPointRibbonViewModelAttribute(string ribbonType) : base(ribbonType)
        {
        }
    }
}