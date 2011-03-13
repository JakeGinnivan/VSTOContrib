using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using Office.Contrib.Extensions;
using CustomTaskPane = Microsoft.Office.Tools.CustomTaskPane;

namespace Office.Contrib.RibbonFactory
{
    internal class ViewModelResolver<TRibbonTypes> : IDisposable where TRibbonTypes : struct 
    {
        private readonly Func<Type, IRibbonViewModel> _ribbonFactory;
        private readonly CustomTaskPaneCollection _customTaskPanes;
        private readonly IViewProvider<TRibbonTypes> _viewProvider;

        /// <summary>
        /// Used when a new explorer or inspector is created to lookup the appropriate viewmodel type
        /// </summary>
        private readonly Dictionary<TRibbonTypes, Type> _ribbonTypeLookup =
            new Dictionary<TRibbonTypes, Type>();
        /// <summary>
        /// Internal lookup for Context instances to view model lookups
        /// </summary>
        private readonly Dictionary<object, IRibbonViewModel> _viewModelInstances = new Dictionary<object, IRibbonViewModel>();
        private readonly Dictionary<object, Queue<CustomTaskPane>> _taskPanesToCleanup = new Dictionary<object, Queue<CustomTaskPane>>();
        private readonly Dictionary<TRibbonTypes, IRibbonUI> _ribbonUiLookup = new Dictionary<TRibbonTypes, IRibbonUI>();
        /// <summary>
        /// Looks up ViewModelType, callback method name, control id, controlId used to invalidate :)
        /// </summary>
        private readonly Dictionary<Type, List<KeyValuePair<string,string>>> _notifyChangeTargetLookup =
            new Dictionary<Type, List<KeyValuePair<string, string>>>();

        private readonly RibbonViewModelHelper _ribbonViewModelHelper;

        public ViewModelResolver(
            IEnumerable<Type> viewModelType, 
            Func<Type, IRibbonViewModel> ribbonFactory, 
            RibbonViewModelHelper ribbonViewModelHelper,
            CustomTaskPaneCollection customTaskPanes,
            IViewProvider<TRibbonTypes> viewProvider)
        {
            _ribbonFactory = ribbonFactory;
            _ribbonViewModelHelper = ribbonViewModelHelper;
            _customTaskPanes = customTaskPanes;
            _viewProvider = viewProvider;

            foreach (var ribbonType in viewModelType)
            {
                CreateRibbonTypeToViewModelTypeLookup(ribbonType);
            }

            _viewProvider.NewView += ViewProviderNewView;
            _viewProvider.ViewClosed += ViewProviderViewClosed;
        }

        void ViewProviderViewClosed(object sender, ViewClosedEventArgs e)
        {
            var views = e.AllViews.ToList();
            foreach (var view in _viewModelInstances.Keys.Where(views.DoesNotContain))
            {
                //Found the inspector that has closed, cleanup viewmodel
                CleanupViewModel(view);
                _viewProvider.CleanupReferencesTo(view);
                view.ReleaseComObject();
            }
        }

        void  ViewProviderNewView(object sender, NewViewEventArgs<TRibbonTypes>  e)
        {
            if (!_ribbonTypeLookup.ContainsKey(e.RibbonType)) return;

            _viewModelInstances.Add(e.ViewInstance, BuildViewModel(e.RibbonType, e.ViewInstance));
            e.Handled = true;
        }

        private void CreateRibbonTypeToViewModelTypeLookup(Type ribbonViewModel)
        {
            foreach (var value in _ribbonViewModelHelper.GetRibbonTypesFor<TRibbonTypes>(ribbonViewModel))
            {
                if (_ribbonTypeLookup.ContainsKey(value))
                    throw new InvalidOperationException("You cannot have two view models which are registered for the same ribbon type");
                _ribbonTypeLookup.Add(value, ribbonViewModel);
            }
        }

        public IRibbonViewModel ResolveInstanceFor(object context)
        {
            return _viewModelInstances[context];
        }

        public void RibbonLoaded(TRibbonTypes currentlyLoadingRibbon, IRibbonUI ribbonUi)
        {
            _ribbonUiLookup.Add(currentlyLoadingRibbon, ribbonUi);

            var viewModelType = _ribbonTypeLookup[currentlyLoadingRibbon];
            foreach (var viewModel in _viewModelInstances.Values
                .Where(viewModel => viewModel.GetType() == viewModelType && viewModel.RibbonUi == null))
            {
                viewModel.RibbonUi = ribbonUi;
            }
        }

        private IRibbonViewModel BuildViewModel(TRibbonTypes ribbonType, object context)
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
            _viewProvider.Dispose();
        }

        public void RegisterCallbackControl(TRibbonTypes ribbonType, string controlCallback, string ribbonControl)
        {
            var type = _ribbonTypeLookup[ribbonType];
            if (!_notifyChangeTargetLookup.ContainsKey(type))
                _notifyChangeTargetLookup.Add(type, new List<KeyValuePair<string, string>>());

            _notifyChangeTargetLookup[type].Add(new KeyValuePair<string, string>(controlCallback, ribbonControl));
        }
    }
}
