using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using NSubstitute;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.Tests.RibbonFactory.TestStubs;
using Xunit;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class the_ribbon_factory : IDisposable
    {
        readonly IViewProvider<TestRibbonTypes> viewProvider;
        readonly TestRibbonFactory ribbonFactoryUnderTest;
        readonly TestViewModelFactory viewModelFactory;

        public the_ribbon_factory()
        {
            viewProvider = Substitute.For<IViewProvider<TestRibbonTypes>>();
            viewModelFactory = new TestViewModelFactory();
            ribbonFactoryUnderTest = new TestRibbonFactory(
                viewModelFactory,
                new Lazy<CustomTaskPaneCollection>(() => Substitute.For<CustomTaskPaneCollection>()),
                viewProvider,
                new TestContextProvider(),
                Assembly.GetExecutingAssembly());
        }

        [Fact]
        public void cannot_create_multiple_instances()
        {
            Assert.Throws<InvalidOperationException>(() => new TestRibbonFactory(
                new TestViewModelFactory(), 
                new Lazy<CustomTaskPaneCollection>(() => Substitute.For<CustomTaskPaneCollection>()),
                viewProvider, new TestContextProvider()));
        }

        [Fact]
        public void cannot_initialise_twice()
        {
            ribbonFactoryUnderTest.SetApplication(null, AddInBaseFactory.Create());

            Assert.Throws<InvalidOperationException>(() => ribbonFactoryUnderTest.SetApplication(null, AddInBaseFactory.Create()));
        }

        [Fact]
        public void default_constructor_uses_default_view_model_locator()
        {
            Assert.IsType<DefaultViewLocationStrategy>(ribbonFactoryUnderTest.LocateViewStrategy);
        }

        [Fact]
        public void initialise_throws_when_no_assemblies_specified_to_scan()
        {
            Assert.Throws<InvalidOperationException>(() => new TestRibbonFactory(new TestViewModelFactory(), 
                new Lazy<CustomTaskPaneCollection>(() => Substitute.For<CustomTaskPaneCollection>()),
                viewProvider, new TestContextProvider()));
        }

        [Fact]
        public void resolves_associated_view_for_viewmodel()
        {
            var customUI1 = ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());
            var customUI2 = ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType2.GetEnumDescription());
            Assert.Contains("view1", customUI1);
            Assert.Contains("view2", customUI2);
        }

        [Fact]
        public void ribbon_xml_callbacks_modified_to_ribbon_factory_callbacks_for_toggle_button()
        {
            // arrange
            // act
            var processedRibbon = ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());

            // assert
            Assert.Contains("onAction=\"PressedOnAction\"", processedRibbon);
            Assert.Contains("getPressed=\"GetPressed\"", processedRibbon);
        }

        [Fact]
        public void ribbon_xml_callbacks_modified_to_ribbon_factory_callbacks_for_button()
        {
            // arrange

            // act
            var processedRibbon = ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());

            // assert
            Assert.Contains("onAction=\"OnAction\"", processedRibbon);
            Assert.Contains("getEnabled=\"GetEnabled\"", processedRibbon);
        }

        [Fact]
        public void toggle_button_is_bound_to_property_get()
        {
            // arrange
            ribbonFactoryUnderTest.SetApplication(null, AddInBaseFactory.Create());
            var processedRibbon = ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow { Context = new TestWindowContext() };
            viewProvider.NewView += Raise.EventWith(viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                viewInstance, viewInstance.Context, TestRibbonTypes.RibbonType1));
            viewModelFactory.ViewModels.Single().PanelShown = true;
            var toggleButtonTag = GetTag(processedRibbon, "testTogglePanelButton");

            // act
            var ribbonControl = GetRibbonControl("testTogglePanelButton", toggleButtonTag, viewInstance);
            var isPressed = ribbonFactoryUnderTest.GetPressed(ribbonControl);

            // assert
            Assert.True(isPressed);
        }

        [Fact]
        public void toggle_button_is_bound_to_property_set()
        {
            // arrange
            ribbonFactoryUnderTest.SetApplication(null, AddInBaseFactory.Create());
            var processedRibbon = ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow { Context = new TestWindowContext() };
            viewProvider.NewView += Raise.EventWith(viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                viewInstance, viewInstance.Context, TestRibbonTypes.RibbonType1));
            viewModelFactory.ViewModels.Single().PanelShown = true;
            var toggleButtonTag = GetTag(processedRibbon, "testTogglePanelButton");

            // act
            var ribbonControl = GetRibbonControl("testTogglePanelButton", toggleButtonTag, viewInstance);
            ribbonFactoryUnderTest.PressedOnAction(ribbonControl, false);

            // assert
            Assert.False(viewModelFactory.ViewModels.Single().PanelShown);
        }

        [Fact]
        public void toggle_button_is_bound_to_property_listens_to_property_changed_events()
        {
            // arrange
            //Open new view to create a viewmodel for view
            ribbonFactoryUnderTest.SetApplication(null, AddInBaseFactory.Create());
            var viewInstance = new TestWindow { Context = new TestWindowContext() };
            viewProvider.NewView += Raise.EventWith(viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                viewInstance, viewInstance.Context, TestRibbonTypes.RibbonType1));
            var ribbon = Substitute.For<IRibbonUI>();
            ribbonFactoryUnderTest.Ribbon_Load(ribbon);

            // act
            viewModelFactory.ViewModels.Single().OnPropertyChanged(new PropertyChangedEventArgs("PanelShown"));

            // assert
            ribbon.Received().InvalidateControl("testTogglePanelButton");
        }

        [Fact]
        public void ribbon_xml_getenabled_can_bind_to_method()
        {
            // arrange
            ribbonFactoryUnderTest.SetApplication(null, AddInBaseFactory.Create());
            var processedRibbon = ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow
                                   {
                                       Context = new TestWindowContext()
                                   };
            viewProvider.NewView += Raise.EventWith(viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                viewInstance, viewInstance.Context, TestRibbonTypes.RibbonType1));
            viewModelFactory.ViewModels.Single().PanelShown = true;
            var buttonTag = GetTag(processedRibbon, "actionButton");

            // act
            var ribbonControl = GetRibbonControl("actionButton", buttonTag, viewInstance);
            var isEnabled = ribbonFactoryUnderTest.GetEnabled(ribbonControl);

            // assert
            Assert.True(isEnabled);
        }

        [Fact]
        public void ribbon_factory_calls_back_to_correct_view_model()
        {
            // arrange
            ribbonFactoryUnderTest.ClearCurrent();
            var ribbonFactory = new TestRibbonFactory(
                viewModelFactory,
                new Lazy<CustomTaskPaneCollection>(() => Substitute.For<CustomTaskPaneCollection>()),
                viewProvider,
                new TestContextProvider(), Assembly.GetExecutingAssembly());
            ribbonFactory.SetApplication(null, AddInBaseFactory.Create());
            var processedRibbon = ribbonFactory.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow { Context = new TestWindowContext() };
            var view2Instance = new TestWindow { Context = new TestWindowContext() };
            viewProvider.NewView += Raise.EventWith(viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                viewInstance, viewInstance.Context, TestRibbonTypes.RibbonType1));
            viewProvider.NewView += Raise.EventWith(viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                view2Instance, view2Instance.Context, TestRibbonTypes.RibbonType1));
            var buttonTag = GetTag(processedRibbon, "testTogglePanelButton");

            // act
            viewModelFactory.ViewModels[1].PanelShown = true;
            var ribbonControl = GetRibbonControl("testTogglePanelButton", buttonTag, viewInstance);
            var ribbon2Control = GetRibbonControl("testTogglePanelButton", buttonTag, view2Instance);
            var isPressed = ribbonFactory.GetPressed(ribbonControl);
            var is2Pressed = ribbonFactory.GetPressed(ribbon2Control);

            // assert
            Assert.False(isPressed);
            Assert.True(is2Pressed);
        }

        [Fact]
        public void new_window_with_same_context_does_not_create_new_viewmodel()
        {
            // arrange
            ribbonFactoryUnderTest.ClearCurrent();
            var ribbonFactory = new TestRibbonFactory(
                viewModelFactory,
                new Lazy<CustomTaskPaneCollection>(() => Substitute.For<CustomTaskPaneCollection>()),
                viewProvider, new TestContextProvider(), Assembly.GetExecutingAssembly());
            ribbonFactory.SetApplication(null, AddInBaseFactory.Create());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow { Context = new TestWindowContext() };
            var view2Instance = new TestWindow { Context = new TestWindowContext() };

            // act

            viewProvider.NewView += Raise.EventWith(viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                                                                        viewInstance, viewInstance.Context,
                                                                        TestRibbonTypes.RibbonType1));
            viewProvider.NewView += Raise.EventWith(viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                                                                        view2Instance, viewInstance.Context,
                                                                        TestRibbonTypes.RibbonType1));

            // assert
            Assert.Equal(1, viewModelFactory.ViewModels.Count);
        }

        [Fact]
        public void new_window_with_different_context_does_not_create_new_viewmodel()
        {
            // arrange
            ribbonFactoryUnderTest.ClearCurrent();
            var ribbonFactory = new TestRibbonFactory(viewModelFactory,
                new Lazy<CustomTaskPaneCollection>(() => Substitute.For<CustomTaskPaneCollection>()),
                viewProvider, new TestContextProvider(), Assembly.GetExecutingAssembly());
            ribbonFactory.SetApplication(null, AddInBaseFactory.Create());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow { Context = new TestWindowContext() };
            var view2Instance = new TestWindow { Context = new TestWindowContext() };

            // act

            viewProvider.NewView += Raise.EventWith(viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                                                                        viewInstance, viewInstance.Context,
                                                                        TestRibbonTypes.RibbonType1));
            viewProvider.NewView += Raise.EventWith(viewProvider, new NewViewEventArgs<TestRibbonTypes>(
                                                                        view2Instance, view2Instance.Context,
                                                                        TestRibbonTypes.RibbonType1));

            // assert
            Assert.Equal(2, viewModelFactory.ViewModels.Count);
        }

        private static string GetTag(string ribbonXml, string controlId)
        {
            var tagExpression = new Regex("\\<.*? id=\\\"" + controlId + "\\\".*?tag=\\\"(.*?)\\\"");
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
            ribbonFactoryUnderTest.ClearCurrent();
        }
    }
}
