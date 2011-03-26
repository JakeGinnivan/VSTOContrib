using System;
using Microsoft.Office.Core;
using Office.Contrib.RibbonFactory;
using Office.Contrib.RibbonFactory.Interfaces;

namespace Office.Contrib.Tests.RibbonFactory.TestStubs
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