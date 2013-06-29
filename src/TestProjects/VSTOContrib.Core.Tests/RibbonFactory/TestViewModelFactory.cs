using System;
using System.Collections.Generic;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.Tests.RibbonFactory.TestStubs;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class TestViewModelFactory : IViewModelFactory
    {
        readonly List<TestRibbonViewModel> viewModels = new List<TestRibbonViewModel>();

        public List<TestRibbonViewModel> ViewModels
        {
            get { return viewModels; }
        }

        public IRibbonViewModel Resolve(Type viewModelType)
        {
            var testRibbon = (TestRibbonViewModel)Activator.CreateInstance(viewModelType);
            ViewModels.Add(testRibbon);
            return testRibbon;
        }

        public void Release(IRibbonViewModel viewModelInstance)
        {
        }
    }
}