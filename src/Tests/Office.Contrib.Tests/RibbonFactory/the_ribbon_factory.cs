using System;
using System.Reflection;
using Microsoft.Office.Tools;
using NSubstitute;
using Office.Contrib.RibbonFactory;
using Office.Contrib.Tests.RibbonFactory.TestStubs;
using Xunit;

namespace Office.Contrib.Tests.RibbonFactory
{
    public class the_ribbon_factory : IDisposable
    {
        private readonly IViewProvider<TestRibbonTypes> _viewProvider;
        private readonly TestRibbonFactory _ribbonFactoryUnderTest;

        public the_ribbon_factory()
        {
            _viewProvider = Substitute.For<IViewProvider<TestRibbonTypes>>();
            _ribbonFactoryUnderTest = new TestRibbonFactory(_viewProvider);
        }

        [Fact]
        public void cannot_create_multiple_instances()
        {
            Assert.Throws<InvalidOperationException>(() => new TestRibbonFactory(_viewProvider));
        }

        [Fact]
        public void cannot_initialise_twice()
        {
            _ribbonFactoryUnderTest.InitialiseFactory(
                t => (IRibbonViewModel) Activator.CreateInstance(t),
                new CustomTaskPaneCollection(),
                Assembly.GetExecutingAssembly());

            Assert.Throws<InvalidOperationException>(() => _ribbonFactoryUnderTest.InitialiseFactory(
                t=>(IRibbonViewModel)Activator.CreateInstance(t),
                new CustomTaskPaneCollection(),
                Assembly.GetExecutingAssembly()));            
        }

        [Fact]
        public void default_constructor_uses_default_view_model_locator()
        {
            Assert.IsType<DefaultViewLocationStrategy>(_ribbonFactoryUnderTest.LocateViewStrategy);
        }

        [Fact]
        public void initialise_throws_when_no_assemblies_specified_to_scan()
        {
            Assert.Throws<InvalidOperationException>(()=>_ribbonFactoryUnderTest.InitialiseFactory(
                t => (IRibbonViewModel)Activator.CreateInstance(t),
                new CustomTaskPaneCollection()));
        }

        [Fact]
        public void resolves_associated_view_for_viewmodel()
        {
            _ribbonFactoryUnderTest.InitialiseFactory(
                t => (IRibbonViewModel) Activator.CreateInstance(t),
                new CustomTaskPaneCollection(),
                Assembly.GetExecutingAssembly());

            var customUI1 = _ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());
            var customUI2 = _ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType2.GetEnumDescription());
            Assert.Contains("view1", customUI1);
            Assert.Contains("view2", customUI2);
        }

        public void Dispose()
        {
            _ribbonFactoryUnderTest.ClearCurrent();
        }
    }
}
