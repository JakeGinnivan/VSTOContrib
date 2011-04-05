using Microsoft.Office.Core;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Core.Tests.RibbonFactory.TestStubs
{
    [RibbonViewModel(TestRibbonTypes.RibbonType2 | TestRibbonTypes.RibbonType3)]
    public class TestRibbonViewModel2 : IRibbonViewModel
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