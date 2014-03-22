using VSTOContrib.Core.RibbonFactory;

namespace VSTOContrib.Outlook.RibbonFactory
{
    /// <summary>
    /// Meta data about the Outlook ribbon view model
    /// </summary>
    public class OutlookRibbonViewModelAttribute : RibbonViewModelAttribute
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        public OutlookRibbonViewModelAttribute(OutlookRibbonType type) : base(type)
        {
        }

        public OutlookRibbonViewModelAttribute(string ribbonType) : base(ribbonType)
        {
        }
    }
}