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

        /// <summary>
        /// The type of Inspector or Explorer that the ribbon should be displayed for.
        /// </summary>
        /// <value>The ribbon type.</value>
        public new OutlookRibbonType Type
        {
            get { return (OutlookRibbonType)base.Type; }
        }
    }
}