using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using Office.Utility.Extensions;
using CustomTaskPane = Microsoft.Office.Tools.CustomTaskPane;

namespace Outlook.Utility.RibbonFactory
{
    internal class ViewModelResolver : IDisposable
    {
        private readonly Func<Type, IRibbonViewModel> _ribbonFactory;
        private readonly CustomTaskPaneCollection _customTaskPanes;
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
        private readonly Dictionary<object, Queue<CustomTaskPane>> _taskPanesToCleanup = new Dictionary<object, Queue<CustomTaskPane>>();
        private readonly Dictionary<RibbonType, IRibbonUI> _ribbonUiLookup = new Dictionary<RibbonType, IRibbonUI>();
        /// <summary>
        /// Looks up ViewModelType, callback method name, control id, controlId used to invalidate :)
        /// </summary>
        private readonly Dictionary<Type, List<KeyValuePair<string,string>>> _notifyChangeTargetLookup =
            new Dictionary<Type, List<KeyValuePair<string, string>>>();

        public ViewModelResolver(
            IEnumerable<Type> viewModelType, 
            Func<Type, IRibbonViewModel> ribbonFactory, 
            _Application outlookApplication,
            CustomTaskPaneCollection customTaskPanes)
        {
            _ribbonFactory = ribbonFactory;
            _customTaskPanes = customTaskPanes;
            _explorers = outlookApplication.Explorers;
            _inspectors = outlookApplication.Inspectors;
            foreach (var ribbonType in viewModelType)
            {
                CreateRibbonTypeToViewModelTypeLookup(ribbonType);
            }
            RegisterExplorers();
            RegisterInspectors();
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
            RegisterCustomTaskPanes(ribbonViewModel, context);
            ListenForINotifyPropertyChanged(ribbonViewModel);

            if (_ribbonUiLookup.ContainsKey(ribbonType))
                ribbonViewModel.RibbonUi = _ribbonUiLookup[ribbonType];

            return ribbonViewModel;
        }

        private void ListenForINotifyPropertyChanged(IRibbonViewModel ribbonViewModel)
        {
            var notifiesOfPropertyChanged = ribbonViewModel as INotifyPropertyChanged;
            if (notifiesOfPropertyChanged != null)
            {
                notifiesOfPropertyChanged.PropertyChanged += NotifiesOfPropertyChangedPropertyChanged;                
            }
        }

        void NotifiesOfPropertyChangedPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            var viewModel = (IRibbonViewModel) sender;
            var senderType = sender.GetType();

            foreach (var invalidatedControl in
                _notifyChangeTargetLookup[senderType]
                    .Where(property => property.Key == e.PropertyName)
                    .Select(pair => pair.Value)
                    .Distinct()
                    .Where(invalidatedControl => viewModel.RibbonUi != null))
            {
                viewModel.RibbonUi.InvalidateControl(invalidatedControl);
            }
        }

        private void RegisterCustomTaskPanes(IRibbonViewModel ribbonViewModel, object context)
        {
            var registersCustomTaskPanes = ribbonViewModel as IRegisterCustomTaskPane;
            if (registersCustomTaskPanes!= null)
            {
                registersCustomTaskPanes.RegisterTaskPanes(
                    (control, title) =>
                        {
                            var taskPane = _customTaskPanes.Add(control, title, context);
                            if (!_taskPanesToCleanup.ContainsKey(context))
                                _taskPanesToCleanup.Add(context, new Queue<CustomTaskPane>());

                            _taskPanesToCleanup[context].Enqueue(taskPane);
                            return taskPane;
                        });
            }
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
            if (_taskPanesToCleanup.ContainsKey(context))
            {
                while (_taskPanesToCleanup[context].Count > 0)
                {
                    _taskPanesToCleanup[context].Dequeue().Dispose();
                }
            }
            var viewModelInstance = _viewModelInstances[context];
            var notifyOfPropertyChanged = viewModelInstance as INotifyPropertyChanged;
            if (notifyOfPropertyChanged != null)
                notifyOfPropertyChanged.PropertyChanged -= NotifiesOfPropertyChangedPropertyChanged;

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

        public void RegisterCallbackControl(RibbonType ribbonType, string controlCallback, string ribbonControl)
        {
            var type = _ribbonTypeLookup[ribbonType];
            if (!_notifyChangeTargetLookup.ContainsKey(type))
                _notifyChangeTargetLookup.Add(type, new List<KeyValuePair<string, string>>());

            _notifyChangeTargetLookup[type].Add(new KeyValuePair<string, string>(controlCallback, ribbonControl));
        }
    }
}
