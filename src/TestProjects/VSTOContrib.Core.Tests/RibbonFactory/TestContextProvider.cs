using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.Tests.RibbonFactory.TestStubs;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class TestContextProvider : IViewContextProvider
    {
        public object GetContextForView(object view)
        {
            return ((TestView)view).Context;
        }

        public string GetRibbonTypeForView(object view)
        {
            return TestRibbonTypes.RibbonType1.GetEnumDescription();
        }
    }
}