using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Office.Utility.Extensions;

namespace Outlook.Utility.RibbonFactory
{
    internal class ViewModelResolver : IDisposable
    {
        private readonly Func<Type, IRibbonViewModel> _ribbonFactory;
        private readonly Explorers _explorers;
        private readonly Inspectors _inspectors;

        /// <summary>
        /// Used when a new explorer or inspector is created to lookup the appropriate viewmodel type
        /// </summary>
        private readonly Dictionary<RibbonType, Type> _ribbonTypeLookup =
            new Dictionary<RibbonType, Type>();
        /// <summary>
        /// Internal lookup for Context instances to view model lookups
        /// </summary>
        private readonly Dictionary<object, IRibbonViewModel> _viewModelInstances = new Dictionary<object, IRibbonViewModel>();
        private readonly Dictionary<RibbonType, IRibbonUI> _ribbonUiLookup = new Dictionary<RibbonType, IRibbonUI>();

        public ViewModelResolver(IEnumerable<Type> viewModelType, Func<Type, IRibbonViewModel> ribbonFactory, Application outlookApplication)
        {
            _ribbonFactory = ribbonFactory;
            _explorers = outlookApplication.Explorers;
            _inspectors = outlookApplication.Inspectors;
            RegisterExplorers();
            RegisterInspectors();
            foreach (var ribbonType in viewModelType)
            {
                CreateRibbonTypeToViewModelTypeLookup(ribbonType);                
            }
        }

        private void CreateRibbonTypeToViewModelTypeLookup(Type ribbonViewModel)
        {
            foreach (var value in RibbonViewModelHelper.GetRibbonTypesFor(ribbonViewModel))
            {
                _ribbonTypeLookup.Add(value, ribbonViewModel);
            }
        }

        public IRibbonViewModel ResolveInstanceFor(object context)
        {
            return _viewModelInstances[context];
        }

        public void RibbonLoaded(RibbonType currentlyLoadingRibbon, IRibbonUI ribbonUi)
        {
            _ribbonUiLookup.Add(currentlyLoadingRibbon, ribbonUi);

            var viewModelType = _ribbonTypeLookup[currentlyLoadingRibbon];
            foreach (var viewModel in _viewModelInstances.Values
                .Where(viewModel => viewModel.GetType() == viewModelType && viewModel.RibbonUi == null))
            {
                viewModel.RibbonUi = ribbonUi;
            }
        }

        private IRibbonViewModel BuildViewModel(RibbonType ribbonType, object context)
        {
            var viewModelType = _ribbonTypeLookup[ribbonType];
            var ribbonViewModel = _ribbonFactory(viewModelType);
            ribbonViewModel.Displayed(context);
            if (_ribbonUiLookup.ContainsKey(ribbonType))
                ribbonViewModel.RibbonUi = _ribbonUiLookup[ribbonType];

            return ribbonViewModel;
        }

        private void RegisterExplorers()
        {
            _explorers.NewExplorer += NewExplorer;

            foreach (Explorer explorer in _explorers)
                NewExplorer(explorer);
        }

        private void RegisterInspectors()
        {
            _inspectors.NewInspector += NewInspector;

            foreach (Inspector inspector in _inspectors)
                NewInspector(inspector);
        }

        void NewInspector(Inspector inspector)
        {
            var ribbonType = InspectorToRibbonTypeConverter.Convert(inspector);

            if (_ribbonTypeLookup.ContainsKey(ribbonType))
            {
                _viewModelInstances.Add(inspector, BuildViewModel(ribbonType, inspector));
                ((InspectorEvents_10_Event)inspector).Close += InspectorClose;
            }
            else
                inspector.ReleaseComObject();
        }

        void NewExplorer(Explorer explorer)
        {
            if (_ribbonTypeLookup.ContainsKey(RibbonType.OutlookExplorer))
            {
                _viewModelInstances.Add(explorer, BuildViewModel(RibbonType.OutlookExplorer, explorer));
                ((ExplorerEvents_10_Event)explorer).Close += ExplorerClose;
                _explorers.Add(explorer);
            }
            else
                explorer.ReleaseComObject();
        }

        void ExplorerClose()
        {
            var explorers = _explorers.Cast<Explorer>().ToList();
            foreach (var explorer in _viewModelInstances.Keys.OfType<Explorer>().Where(explorers.DoesNotContain))
            {
                //Found the explorer that has closed, cleanup viewmodel
                CleanupViewModel(explorer);
                ((ExplorerEvents_10_Event)explorer).Close -= ExplorerClose;
                explorer.ReleaseComObject();
                break;
            }
            foreach (var explorer in explorers)
            {
                explorer.ReleaseComObject();
            }
        }

        void InspectorClose()
        {
            var inspectors = _inspectors.Cast<Inspector>().ToList();
            foreach (var inspector in _viewModelInstances.Keys.OfType<Inspector>().Where(inspectors.DoesNotContain))
            {
                //Found the inspector that has closed, cleanup viewmodel
                CleanupViewModel(inspector);
                ((InspectorEvents_10_Event)inspector).Close -= ExplorerClose;
                inspector.ReleaseComObject();
                break;
            }
            foreach (var inspector in inspectors)
            {
                inspector.ReleaseComObject();
            }
        }

        private void CleanupViewModel(object context)
        {
            var viewModelInstance = _viewModelInstances[context];
            var disposible = viewModelInstance as IDisposable;
            if (disposible != null) disposible.Dispose();
            viewModelInstance.Cleanup();
            _viewModelInstances.Remove(context);
        }

        public void Dispose()
        {
            _explorers.NewExplorer -= NewExplorer;
            _inspectors.NewInspector -= NewInspector;
            _explorers.ReleaseComObject();
            _inspectors.ReleaseComObject();
        }
    }
}
