﻿using System;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using NSubstitute;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.Tests.RibbonFactory.TestAddin;
using VSTOContrib.Core.Tests.RibbonFactory.TestStubs;
using Xunit;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class the_ribbon_factory : IDisposable
    {
        readonly IViewProvider viewProvider;
        readonly TestRibbonFactory ribbonFactoryUnderTest;
        readonly TestViewModelFactory viewModelFactory;
        readonly TestAddInBase testAddInBase;

        public the_ribbon_factory()
        {
            viewProvider = Substitute.For<IViewProvider>();
            viewModelFactory = new TestViewModelFactory();
            testAddInBase = AddInBaseFactory.Create();
            if (Core.RibbonFactory.RibbonFactory.Current == null)
            {
                ribbonFactoryUnderTest = new TestRibbonFactory(
                testAddInBase,
                viewProvider,
                new TestContextProvider(),
                "Foo",
                Assembly.GetExecutingAssembly())
                {
                    ViewModelFactory = viewModelFactory
                };
            }
            else
            {
                ribbonFactoryUnderTest = (TestRibbonFactory)Core.RibbonFactory.RibbonFactory.Current;
            }
        }

        [Fact]
        public void cannot_create_multiple_instances()
        {
            Assert.Throws<InvalidOperationException>(() => new TestRibbonFactory(
                testAddInBase, 
                viewProvider,
                new TestContextProvider(),
                "Foo",
                Assembly.GetExecutingAssembly()));
        }

        [Fact]
        public void default_constructor_uses_default_view_model_locator()
        {
            Assert.IsType<DefaultViewLocationStrategy>(ribbonFactoryUnderTest.LocateViewStrategy);
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
            testAddInBase.TestAddin.OnStartup();
            var processedRibbon = ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow { Context = new TestWindowContext() };
            viewProvider.NewView += Raise.EventWith(viewProvider, new NewViewEventArgs(
                viewInstance, viewInstance.Context, TestRibbonTypes.RibbonType1.GetEnumDescription()));
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
            testAddInBase.TestAddin.OnStartup();
            var processedRibbon = ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow { Context = new TestWindowContext() };
            viewProvider.NewView += Raise.EventWith(viewProvider, new NewViewEventArgs(
                viewInstance, viewInstance.Context, TestRibbonTypes.RibbonType1.GetEnumDescription()));
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
            testAddInBase.TestAddin.OnStartup();
            var viewInstance = new TestWindow { Context = new TestWindowContext() };
            viewProvider.NewView += Raise.EventWith(viewProvider, new NewViewEventArgs(
                viewInstance, viewInstance.Context, TestRibbonTypes.RibbonType1.GetEnumDescription()));
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
            testAddInBase.TestAddin.OnStartup();
            var processedRibbon = ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow
                                   {
                                       Context = new TestWindowContext()
                                   };
            viewProvider.NewView += Raise.EventWith(viewProvider, new NewViewEventArgs(
                viewInstance, viewInstance.Context, TestRibbonTypes.RibbonType1.GetEnumDescription()));
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
            testAddInBase.TestAddin.OnStartup();
            var processedRibbon = ribbonFactoryUnderTest.GetCustomUI(TestRibbonTypes.RibbonType1.GetEnumDescription());
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow { Context = new TestWindowContext() };
            var view2Instance = new TestWindow { Context = new TestWindowContext() };
            viewProvider.NewView += Raise.EventWith(viewProvider, new NewViewEventArgs(
                viewInstance, viewInstance.Context, TestRibbonTypes.RibbonType1.GetEnumDescription()));
            viewProvider.NewView += Raise.EventWith(viewProvider, new NewViewEventArgs(
                view2Instance, view2Instance.Context, TestRibbonTypes.RibbonType1.GetEnumDescription()));
            var buttonTag = GetTag(processedRibbon, "testTogglePanelButton");

            // act
            viewModelFactory.ViewModels[1].PanelShown = true;
            var ribbonControl = GetRibbonControl("testTogglePanelButton", buttonTag, viewInstance);
            var ribbon2Control = GetRibbonControl("testTogglePanelButton", buttonTag, view2Instance);
            var isPressed = ribbonFactoryUnderTest.GetPressed(ribbonControl);
            var is2Pressed = ribbonFactoryUnderTest.GetPressed(ribbon2Control);

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
                testAddInBase,
                viewProvider, new TestContextProvider(), "Foo", Assembly.GetExecutingAssembly());
            testAddInBase.TestAddin.OnStartup();
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow { Context = new TestWindowContext() };
            var view2Instance = new TestWindow { Context = new TestWindowContext() };

            // act

            var viewEventArgs = new NewViewEventArgs(viewInstance, viewInstance.Context,
                TestRibbonTypes.RibbonType1.GetEnumDescription());
            viewProvider.NewView += Raise.EventWith(viewProvider, viewEventArgs);
            var newViewEventArgs = new NewViewEventArgs(view2Instance, viewInstance.Context,
                TestRibbonTypes.RibbonType1.GetEnumDescription());
            viewProvider.NewView += Raise.EventWith(viewProvider, newViewEventArgs);

            // assert
            Assert.Equal(1, viewModelFactory.ViewModels.Count);
        }

        [Fact]
        public void new_window_with_different_context_does_not_create_new_viewmodel()
        {
            // arrange
            ribbonFactoryUnderTest.ClearCurrent();
            testAddInBase.TestAddin.OnStartup();
            //Open new view to create a viewmodel for view
            var viewInstance = new TestWindow { Context = new TestWindowContext() };
            var view2Instance = new TestWindow { Context = new TestWindowContext() };

            // act
            var newViewEventArgs = new NewViewEventArgs(viewInstance, viewInstance.Context,
                TestRibbonTypes.RibbonType1.GetEnumDescription());
            var viewEventArgs = new NewViewEventArgs(view2Instance, view2Instance.Context,
                TestRibbonTypes.RibbonType1.GetEnumDescription());
            viewProvider.NewView += Raise.EventWith(viewProvider, newViewEventArgs);
            viewProvider.NewView += Raise.EventWith(viewProvider, viewEventArgs);

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
            ((object)ribbonControl.Context).Returns(view);
            return ribbonControl;
        }

        public void Dispose()
        {
            ribbonFactoryUnderTest.ClearCurrent();
        }
    }
}
