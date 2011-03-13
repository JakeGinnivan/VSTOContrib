using Office.Contrib.RibbonFactory;

namespace Office.Contrib.Tests.RibbonFactory.TestStubs
{
    public class TestRibbonFactory : RibbonFactory<TestRibbonTypes>
    {
        private readonly IViewProvider<TestRibbonTypes> _viewProvider;

        public TestRibbonFactory(IViewProvider<TestRibbonTypes> viewProvider)
        {
            _viewProvider = viewProvider;
        }

        protected override IViewProvider<TestRibbonTypes> ViewProvider()
        {
            return _viewProvider;
        }

        public void ClearCurrent()
        {
            Current = null;
        }
    }
}