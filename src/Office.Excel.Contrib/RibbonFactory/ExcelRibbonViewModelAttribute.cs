using Office.Contrib.RibbonFactory;

namespace Office.Excel.Contrib.RibbonFactory
{
    /// <summary>
    /// Meta data about the Outlook ribbon view model
    /// </summary>
    public class ExcelRibbonViewModelAttribute : RibbonViewModelAttribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelRibbonViewModelAttribute"/> class.
        /// </summary>
        public ExcelRibbonViewModelAttribute()
            : base(ExcelRibbonType.ExcelWorkbook)
        {
        }

        /// <summary>
        /// The type of Inspector or Explorer that the ribbon should be displayed for.
        /// </summary>
        /// <value>The ribbon type.</value>
        public new ExcelRibbonType Type
        {
            get { return (ExcelRibbonType)base.Type; }
        }
    }
}