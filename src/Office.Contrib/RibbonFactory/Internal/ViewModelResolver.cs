using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Microsoft.Office.Core;
using Office.Contrib.RibbonFactory.Interfaces;
using Office.Contrib.RibbonFactory.Interfaces.Internal;

namespace Office.Contrib.RibbonFactory.Internal
{
    internal class ViewModelResolver<TRibbonTypes> : IDisposable, IViewModelResolver<TRibbonTypes> where TRibbonTypes : struct 
    {
        /// <summary>
        /// Used when a new explorer or inspector is created to lookup the appropriate viewmodel type
        /// </summary>
        private readonly Dictionary<TRibbonTypes, Type> _ribbonTypeLookup;
        /// <summary>
        /// Internal lookup for Context instances to view model lookups
        /// </summary>
        private readonly Dictionary<object, IRibbonViewModel> _contextToViewModelLookup;
        private readonly Dictionary<TRibbonTypes, IRibbonUI> _ribbonUiLookup;
        /// <summary>
        /// Looks up ViewModelType, callback method name, control id, controlId used to invalidate :)
        /// </summary>
        private readonly Dictionary<Type, List<KeyValuePair<string,string>>> _notifyChangeTargetLookup;

        private readonly RibbonViewModelHelper _ribbonViewModelHelper;
        private readonly ICustomTaskPaneRegister _customTaskPaneRegister;
        private Func<Type, IRibbonViewModel> _ribbonFactory;
        private IViewProvider<TRibbonTypes> _viewProvider;
        private TRibbonTypes _currentlyLoadingRibbon;
        private IViewContextProvider _viewContextProvider;

        public ViewModelResolver(
            IEnumerable<Type> viewModelType, 
            RibbonViewModelHelper ribbonViewModelHelper,
            ICustomTaskPaneRegister customTaskPaneRegister)
        {
            _notifyChangeTargetLookup = new Dictionary<Type, List<KeyValuePair<string, string>>>();
            _ribbonTypeLookup = new Dictionary<TRibbonTypes, Type>();
            _contextToViewModelLookup = new Dictionary<object, IRibbonViewModel>();
            _ribbonUiLookup = new Dictionary<TRibbonTypes, IRibbonUI>();
            _ribbonViewModelHelper = ribbonViewModelHelper;
            _customTaskPaneRegister = customTaskPaneRegister;

            foreach (var ribbonType in viewModelType)
            {
                CreateRibbonTypeToViewModelTypeLookup(ribbonType);
            }
        }

        public void Initialise(
            Func<Type, IRibbonViewModel> ribbonFactory,
            IViewProvider<TRibbonTypes> viewProvider,
            IViewContextProvider viewContextProvider)
        {
            _viewContextProvider = viewContextProvider;
            _ribbonFactory = ribbonFactory;
            _viewProvider = viewProvider;

            _viewProvider.NewView += ViewProviderNewView;
            _viewProvider.ViewClosed += ViewProviderViewClosed;
        }

        void ViewProviderViewClosed(object sender, ViewClosedEventArgs e)
        {
            CleanupViewModel(e.Context);
            _viewProvider.CleanupReferencesTo(e.View, e.Context);
        }

        void ViewProviderNewView(object sender, NewViewEventArgs<TRibbonTypes> e)
        {
            if (!_ribbonTypeLookup.ContainsKey(e.RibbonType)) return;
            if (_contextToViewModelLookup.ContainsKey(e.ViewContext)) return; //Reuse viewmodels for each context

            _currentlyLoadingRibbon = e.RibbonType;
            _contextToViewModelLookup.Add(e.ViewContext, BuildViewModel(e.RibbonType, e.ViewInstance, e.ViewContext));

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

        public IRibbonViewModel ResolveInstanceFor(object view)
        {
            var context = _viewContextProvider.GetContextForView(view);
            return _contextToViewModelLookup[context];
        }

        public void RibbonLoaded(IRibbonUI ribbonUi)
        {
            _ribbonUiLookup.Add(_currentlyLoadingRibbon, ribbonUi);

            var viewModelType = _ribbonTypeLookup[_currentlyLoadingRibbon];
            foreach (var viewModelLookup in _contextToViewModelLookup.Values
                .Where(viewModel => viewModel.GetType() == viewModelType && viewModel.RibbonUi == null))
            {
                viewModelLookup.RibbonUi = ribbonUi;
            }
        }

        private IRibbonViewModel BuildViewModel(TRibbonTypes ribbonType, object viewInstance, object viewContext)
        {
            var viewModelType = _ribbonTypeLookup[ribbonType];
            var ribbonViewModel = _ribbonFactory(viewModelType);
            ribbonViewModel.Displayed(viewContext);
            _customTaskPaneRegister.RegisterCustomTaskPanes(ribbonViewModel, viewInstance);
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

        private void CleanupViewModel(object context)
        {
            //TODO write test around context/view cleanup
            _customTaskPaneRegister.Cleanup(context);
            var viewModelInstance = _contextToViewModelLookup[context];


            var notifyOfPropertyChanged = viewModelInstance as INotifyPropertyChanged;
            if (notifyOfPropertyChanged != null)
                notifyOfPropertyChanged.PropertyChanged -= NotifiesOfPropertyChangedPropertyChanged;

            var disposible = viewModelInstance as IDisposable;
            if (disposible != null) disposible.Dispose();
            viewModelInstance.Cleanup();

            _contextToViewModelLookup.Remove(context);
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

    internal class ViewModelKey
    {
        public object View { get; private set; }
        public object Context { get; private set; }

        public ViewModelKey(object view, object context)
        {
            View = view;
            Context = context;
        }
    }
}
