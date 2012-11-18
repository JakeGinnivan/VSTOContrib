using Microsoft.Office.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Word.RibbonFactory;

namespace WordAddIn1
{
    [WordRibbonViewModel]
    public class DocumentViewModel : OfficeViewModelBase, IRibbonViewModel
    {
        public IRibbonUI RibbonUi { get; set; }

        public void Initialised(object context)
        {
        }

        public void CurrentViewChanged(object currentView)
        {
        }

        public void Cleanup()
        {
        }
    }
}
