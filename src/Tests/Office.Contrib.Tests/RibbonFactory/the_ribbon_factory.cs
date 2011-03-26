using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using NSubstitute;
using Office.Contrib.RibbonFactory;
using Office.Contrib.RibbonFactory.Interfaces;
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
            _ribbonFactoryUnderTest = new TestRibbonFactory(_viewProvider, new TestContextProvider(), Assembly.GetExecutingAssembly());
        }

        [Fact]
        public void cannot_create_multiple_instances()
        {
            Assert.Throws<InvalidOperationException>(() => new TestRibbonFactory(_viewProvider, new TestContextProvider()));
        }

        [Fact]
        public void cannot_initialise_twice()
        {
            _ribbonFactoryUnderTest.InitialiseFactory(
                t => (IRibbonViewModel) Activator.CreateInstance(t),
                new CustomTaskPaneCollection());

            Assert.Throws<InvalidOperationException>(() => _ribbonFactoryUnderTest.InitialiseFactory(
                t=>(IRibbonViewModel)Activator.CreateInstance(t),
                new CustomTaskPaneCollection()));            
        }

        [Fact]
        public void default_constructor_uses_default_view_model_locator()
        {
            Assert.IsType<DefaultViewLocationStrategy>(_ribbonFactoryUnderTest.LocateViewStrategy);
        }

        [Fact]
        public void initialise_throws_when_no_assemblies_specified_to_scan()
        {
            Assert.Throws<InvalidOperationException>(()=>new TestRibbonFactory(_viewProvider, new TestContextProvider()));
        }

        [Fact]
        public void resolves_associated_view_for_viewmodel()
        {
            _ribbonFactoryUnderTest.InitialiseFactory(
                t => (IRibbonViewModel) Activator.CreateInstance(t),
                new CustomTaskPaneCollection());

            var customUI1 = _ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());
            var customUI2 = _ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType2.GetEnumDescription());
            Assert.Contains("view1", customUI1);
            Assert.Contains("view2", customUI2);
        }

        [Fact]
        public void ribbon_xml_callbacks_modified_to_ribbon_factory_callbacks_for_toggle_button()
        {
            // arrange
            _ribbonFactoryUnderTest.InitialiseFactory(
                t => (IRibbonViewModel) Activator.CreateInstance(t),
                new CustomTaskPaneCollection());

            // act
            var processedRibbon = _ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());

            // assert
            Assert.Contains("onAction=\"PressedOnAction\"", processedRibbon);
            Assert.Contains("getPressed=\"GetPressed\"", processedRibbon);
        }

        [Fact]
        public void ribbon_xml_callbacks_modified_to_ribbon_factory_callbacks_for_button()
        {
            // arrange
            _ribbonFactoryUnderTest.InitialiseFactory(
                t => (IRibbonViewModel)Activator.CreateInstance(t),
                new CustomTaskPaneCollection());

            // act
            var processedRibbon = _ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());

            // assert
            Assert.Contains("onAction=\"OnAction\"", processedRibbon);
            Assert.Contains("getEnabled=\"GetEnabled\"", processedRibbon);
        }

        [Fact]
        public void toggle_button_is_bound_to_property_get()
        {
            // arrange
            TestRibbonViewModel viewModel = null;
            _ribbonFactoryUnderTest.InitialiseFactory(
                t => viewModel = (TestRibbonViewModel)Activator.CreateInstance(t),
                new CustomTaskPaneCollection());
            var processedRibbon = _ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow{Context = new TestWindowContext()};
            _viewProvider.NewView += Raise.EventWith(_viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                viewInstance, viewInstance.Context, TestRibbonTypes.RibbonType1));
            viewModel.PanelShown = true;
            var toggleButtonTag = GetTag(processedRibbon, "testTogglePanelButton");

            // act
            var ribbonControl = GetRibbonControl("testTogglePanelButton", toggleButtonTag, viewInstance);
            var isPressed = _ribbonFactoryUnderTest.GetPressed(ribbonControl);

            // assert
            Assert.True(isPressed);
        }

        [Fact]
        public void toggle_button_is_bound_to_property_set()
        {
            // arrange
            TestRibbonViewModel viewModel = null;
            _ribbonFactoryUnderTest.InitialiseFactory(
                t => viewModel = (TestRibbonViewModel)Activator.CreateInstance(t),
                new CustomTaskPaneCollection());
            var processedRibbon = _ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow{ Context = new TestWindowContext()};
            _viewProvider.NewView += Raise.EventWith(_viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                viewInstance, viewInstance.Context, TestRibbonTypes.RibbonType1));
            viewModel.PanelShown = true;
            var toggleButtonTag = GetTag(processedRibbon, "testTogglePanelButton");

            // act
            var ribbonControl = GetRibbonControl("testTogglePanelButton", toggleButtonTag, viewInstance);
            _ribbonFactoryUnderTest.PressedOnAction(ribbonControl, false);

            // assert
            Assert.False(viewModel.PanelShown);
        }

        [Fact]
        public void toggle_button_is_bound_to_property_listens_to_property_changed_events()
        {
            // arrange
            TestRibbonViewModel viewModel = null;
            _ribbonFactoryUnderTest.InitialiseFactory(
                t => viewModel = (TestRibbonViewModel)Activator.CreateInstance(t),
                new CustomTaskPaneCollection());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow{Context = new TestWindowContext()};
            _viewProvider.NewView += Raise.EventWith(_viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                viewInstance, viewInstance.Context, TestRibbonTypes.RibbonType1));
            var ribbon = Substitute.For<IRibbonUI>();
            _ribbonFactoryUnderTest.Ribbon_Load(ribbon);

            // act
            viewModel.OnPropertyChanged(new PropertyChangedEventArgs("PanelShown"));

            // assert
            ribbon.Received().InvalidateControl("testTogglePanelButton");
        }

        [Fact]
        public void ribbon_xml_getenabled_can_bind_to_method()
        {
            // arrange
            TestRibbonViewModel viewModel = null;
            _ribbonFactoryUnderTest.InitialiseFactory(
                t => viewModel = (TestRibbonViewModel)Activator.CreateInstance(t),
                new CustomTaskPaneCollection());
            var processedRibbon = _ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow
                                   {
                                       Context = new TestWindowContext()
                                   };
            _viewProvider.NewView += Raise.EventWith(_viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                viewInstance, viewInstance.Context, TestRibbonTypes.RibbonType1));
            viewModel.PanelShown = true;
            var buttonTag = GetTag(processedRibbon, "actionButton");

            // act
            var ribbonControl = GetRibbonControl("actionButton", buttonTag, viewInstance);
            var isEnabled = _ribbonFactoryUnderTest.GetEnabled(ribbonControl);

            // assert
            Assert.True(isEnabled);
        }

        [Fact]
        public void ribbon_factory_calls_back_to_correct_view_model()
        {
            // arrange
            var viewModels = new List<TestRibbonViewModel>();
            _ribbonFactoryUnderTest.InitialiseFactory(
                t =>
                    {
                        var testRibbon = (TestRibbonViewModel)Activator.CreateInstance(t);
                        viewModels.Add(testRibbon);
                        return testRibbon;
                    },
                new CustomTaskPaneCollection());
            var processedRibbon = _ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow{Context = new TestWindowContext()};
            var view2Instance = new TestWindow { Context = new TestWindowContext() };
            _viewProvider.NewView += Raise.EventWith(_viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                viewInstance, viewInstance.Context, TestRibbonTypes.RibbonType1));
            _viewProvider.NewView += Raise.EventWith(_viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                view2Instance, view2Instance.Context, TestRibbonTypes.RibbonType1));
            var buttonTag = GetTag(processedRibbon, "testTogglePanelButton");

            // act
            viewModels[1].PanelShown = true;
            var ribbonControl = GetRibbonControl("testTogglePanelButton", buttonTag, viewInstance);
            var ribbon2Control = GetRibbonControl("testTogglePanelButton", buttonTag, view2Instance);
            var isPressed = _ribbonFactoryUnderTest.GetPressed(ribbonControl);
            var is2Pressed = _ribbonFactoryUnderTest.GetPressed(ribbon2Control);

            // assert
            Assert.False(isPressed);
            Assert.True(is2Pressed);
        }

        [Fact]
        public void new_window_with_same_context_does_not_create_new_viewmodel()
        {
            // arrange
            var viewModels = new List<TestRibbonViewModel>();
            _ribbonFactoryUnderTest.InitialiseFactory(
                t =>
                    {
                        var testRibbon = (TestRibbonViewModel) Activator.CreateInstance(t);
                        viewModels.Add(testRibbon);
                        return testRibbon;
                    },
                new CustomTaskPaneCollection());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow {Context = new TestWindowContext()};
            var view2Instance = new TestWindow {Context = new TestWindowContext()};

            // act

            _viewProvider.NewView += Raise.EventWith(_viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                                                                        viewInstance, viewInstance.Context,
                                                                        TestRibbonTypes.RibbonType1));
            _viewProvider.NewView += Raise.EventWith(_viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                                                                        view2Instance, viewInstance.Context,
                                                                        TestRibbonTypes.RibbonType1));

            // assert
            Assert.Equal(1, viewModels.Count);
        }

        [Fact]
        public void new_window_with_different_context_does_not_create_new_viewmodel()
        {
            // arrange
            var viewModels = new List<TestRibbonViewModel>();
            _ribbonFactoryUnderTest.InitialiseFactory(
                t =>
                {
                    var testRibbon = (TestRibbonViewModel)Activator.CreateInstance(t);
                    viewModels.Add(testRibbon);
                    return testRibbon;
                },
                new CustomTaskPaneCollection());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow { Context = new TestWindowContext() };
            var view2Instance = new TestWindow { Context = new TestWindowContext() };

            // act

            _viewProvider.NewView += Raise.EventWith(_viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                                                                        viewInstance, viewInstance.Context,
                                                                        TestRibbonTypes.RibbonType1));
            _viewProvider.NewView += Raise.EventWith(_viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                                                                        view2Instance, view2Instance.Context,
                                                                        TestRibbonTypes.RibbonType1));

            // assert
            Assert.Equal(2, viewModels.Count);
        }

        private static string GetTag(string ribbonXml, string controlId)
        {
            var tagExpression = new Regex("\\<.*? id=\\\""+controlId+"\\\".*?tag=\\\"(.*?)\\\"");
            return tagExpression.Match(ribbonXml).Groups[1].Value;
        }

        private static IRibbonControl GetRibbonControl(string id, string tag, object view)
        {
            var ribbonControl = Substitute.For<IRibbonControl>();
            ribbonControl.Id.Returns(id);
            ribbonControl.Tag.Returns(tag);
            ribbonControl.Context.Returns(view);
            return ribbonControl;
        }

        public void Dispose()
        {
            _ribbonFactoryUnderTest.ClearCurrent();
        }
    }

    public class TestContextProvider : IViewContextProvider
    {
        public object GetContextForView(object view)
        {
            return ((TestWindow) view).Context;
        }
    }

    public class TestWindowContext
    {
    }

    public class TestWindow
    {
        public TestWindowContext Context { get; set; }
    }
}
