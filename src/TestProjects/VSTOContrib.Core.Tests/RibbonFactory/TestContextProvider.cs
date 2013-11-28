using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.Tests.RibbonFactory.TestStubs;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class TestContextProvider : IViewContextProvider
    {
        public object GetContextForView(object view)
        {
            return ((TestWindow)view).Context;
        }

        public TRibbonType GetRibbonTypeForView<TRibbonType>(object view)
        {
            return (TRibbonType)(object)TestRibbonTypes.RibbonType1;
        }
    }
}