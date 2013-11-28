using VSTOContrib.Core.RibbonFactory;

namespace VSTOContrib.Excel.RibbonFactory
{
    public class ExcelRibbonViewModelAttribute : RibbonViewModelAttribute
    {
        public ExcelRibbonViewModelAttribute()
            : base(ExcelRibbonType.ExcelWorkbook)
        {
        }

        public ExcelRibbonViewModelAttribute(string ribbonType) : base(ribbonType)
        {
        }
    }
}