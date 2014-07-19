using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.Tests.RibbonFactory.TestStubs;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class TestContextProvider : IViewContextProvider
    {
        public object GetContextForView(OfficeWin32Window view)
        {
            return ((TestView)view.Window).Context;
        }

        public string GetRibbonTypeForView(OfficeWin32Window view)
        {
            return TestRibbonTypes.RibbonType1.GetEnumDescription();
        }
    }
}