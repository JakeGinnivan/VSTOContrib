using Office.Contrib.RibbonFactory;

namespace Office.PowerPoint.Contrib.RibbonFactory
{
    /// <summary>
    /// Meta data about the Outlook ribbon view model
    /// </summary>
    public class PowerPointRibbonViewModelAttribute : RibbonViewModelAttribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PowerPointRibbonViewModelAttribute"/> class.
        /// </summary>
        public PowerPointRibbonViewModelAttribute()
            : base(PowerPointRibbonType.PowerPointPresentation)
        {
        }

        /// <summary>
        /// The type of Inspector or Explorer that the ribbon should be displayed for.
        /// </summary>
        /// <value>The ribbon type.</value>
        public new PowerPointRibbonType Type
        {
            get { return (PowerPointRibbonType)base.Type; }
        }
    }
}