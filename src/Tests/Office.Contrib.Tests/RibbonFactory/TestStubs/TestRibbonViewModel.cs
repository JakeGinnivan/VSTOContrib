using System;
using Microsoft.Office.Core;
using Office.Contrib.RibbonFactory;

namespace Office.Contrib.Tests.RibbonFactory.TestStubs
{
    [RibbonViewModel(TestRibbonTypes.RibbonType1)]
    public class TestRibbonViewModel : IRibbonViewModel
    {
        public IRibbonUI RibbonUi { get; set; }

        public void Displayed(object context)
        {
        }

        public void Cleanup()
        {
        }
    }
}