using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Microsoft.Office.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Interfaces.Internal;

namespace VSTOContrib.Core.RibbonFactory.Internal
{
    internal class ViewModelResolver<TRibbonTypes> : IViewModelResolver<TRibbonTypes> where TRibbonTypes : struct 
    {
        /// <summary>
        /// Used when a new explorer or inspector is created to lookup the appropriate viewmodel type
        /// </summary>
        private readonly Dictionary<TRibbonTypes, Type> ribbonTypeLookup;
        /// <summary>
        /// Internal lookup for Context instances to view model lookups
        /// </summary>
        private readonly Dictionary<object, IRibbonViewModel> contextToViewModelLookup;
        private readonly Dictionary<TRibbonTypes, IRibbonUI> ribbonUiLookup;
        /// <summary>
        /// Looks up ViewModelType, callback method name, control id, controlId used to invalidate :)
        /// </summary>
        private readonly Dictionary<Type, List<KeyValuePair<string,string>>> notifyChangeTargetLookup;

        private readonly RibbonViewModelHelper ribbonViewModelHelper;
        private readonly ICustomTaskPaneRegister customTaskPaneRegister;
        private readonly IViewContextProvider viewContextProvider;
        private readonly Func<Type, IRibbonViewModel> ribbonFactory;
        private IViewProvider<TRibbonTypes> viewProvider;
        private TRibbonTypes currentlyLoadingRibbon;

        public ViewModelResolver(IEnumerable<Type> viewModelType, RibbonViewModelHelper ribbonViewModelHelper, ICustomTaskPaneRegister customTaskPaneRegister, 
            IViewContextProvider viewContextProvider,
            Func<Type, IRibbonViewModel> ribbonFactory
            )
        {
            currentlyLoadingRibbon = (TRibbonTypes)(object)1;
            notifyChangeTargetLookup = new Dictionary<Type, List<KeyValuePair<string, string>>>();
            ribbonTypeLookup = new Dictionary<TRibbonTypes, Type>();
            contextToViewModelLookup = new Dictionary<object, IRibbonViewModel>();
            ribbonUiLookup = new Dictionary<TRibbonTypes, IRibbonUI>();
            this.ribbonViewModelHelper = ribbonViewModelHelper;
            this.customTaskPaneRegister = customTaskPaneRegister;
            this.viewContextProvider = viewContextProvider;
            this.ribbonFactory = ribbonFactory;

            foreach (var ribbonType in viewModelType)
            {
                CreateRibbonTypeToViewModelTypeLookup(ribbonType);
            }
        }

        public void Initialise(
            IViewProvider<TRibbonTypes> viewProvider)
        {
            this.viewProvider = viewProvider;

            this.viewProvider.NewView += ViewProviderNewView;
            this.viewProvider.ViewClosed += ViewProviderViewClosed;
        }

        void ViewProviderViewClosed(object sender, ViewClosedEventArgs e)
        {
            //TODO write test around context/view cleanup
            customTaskPaneRegister.Cleanup(e.View);

            CleanupViewModel(e.Context);
            viewProvider.CleanupReferencesTo(e.View, e.Context);
        }

        void ViewProviderNewView(object sender, NewViewEventArgs<TRibbonTypes> e)
        {
            if (!ribbonTypeLookup.ContainsKey(e.RibbonType)) return;
            if (contextToViewModelLookup.ContainsKey(e.ViewContext))
            {
                //Tell viewmodel there is a new view active
                var ribbonViewModel = contextToViewModelLookup[e.ViewContext];
                ribbonViewModel.CurrentViewChanged(e.ViewInstance);
                customTaskPaneRegister.RegisterCustomTaskPanes(ribbonViewModel, e.ViewInstance);
                return; 
            }

            currentlyLoadingRibbon = e.RibbonType;
            contextToViewModelLookup.Add(e.ViewContext, BuildViewModel(e.RibbonType, e.ViewInstance, e.ViewContext));

            e.Handled = true;
        }

        private void CreateRibbonTypeToViewModelTypeLookup(Type ribbonViewModel)
        {
            foreach (var value in ribbonViewModelHelper.GetRibbonTypesFor<TRibbonTypes>(ribbonViewModel))
            {
                if (ribbonTypeLookup.ContainsKey(value))
                    throw new InvalidOperationException("You cannot have two view models which are registered for the same ribbon type");
                ribbonTypeLookup.Add(value, ribbonViewModel);
            }
        }

        public IRibbonViewModel ResolveInstanceFor(object view)
        {
            var context = viewContextProvider.GetContextForView(view) ?? NullContext.Instance;

            //Sometimes can happen that view provider has not got events to tell us about a new view
            // so we will have to try and create it
            if (!contextToViewModelLookup.ContainsKey(context))
            {
                var ribbonTypeForView = viewContextProvider.GetRibbonTypeForView<TRibbonTypes>(view);
                var newViewEventArgs = new NewViewEventArgs<TRibbonTypes>(view, context, ribbonTypeForView);

                ViewProviderNewView(this, newViewEventArgs);
            }

            return contextToViewModelLookup[context];
        }

        public void RibbonLoaded(IRibbonUI ribbonUi)
        {
            ribbonUiLookup.Add(currentlyLoadingRibbon, ribbonUi);

            if (!ribbonTypeLookup.ContainsKey(currentlyLoadingRibbon))
                return;
            var viewModelType = ribbonTypeLookup[currentlyLoadingRibbon];
            foreach (var viewModelLookup in contextToViewModelLookup.Values
                .Where(viewModel => viewModel.GetType() == viewModelType && viewModel.RibbonUi == null))
            {
                viewModelLookup.RibbonUi = ribbonUi;
            }
        }

        private IRibbonViewModel BuildViewModel(TRibbonTypes ribbonType, object viewInstance, object viewContext)
        {
            var viewModelType = ribbonTypeLookup[ribbonType];
            var ribbonViewModel = ribbonFactory(viewModelType);

            if (ribbonUiLookup.ContainsKey(ribbonType))
                ribbonViewModel.RibbonUi = ribbonUiLookup[ribbonType];

            ListenForINotifyPropertyChanged(ribbonViewModel);
            ribbonViewModel.Initialised(viewContext == NullContext.Instance ? null : viewContext);
            ribbonViewModel.CurrentViewChanged(viewInstance);

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
                notifyChangeTargetLookup[senderType]
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
            if (!contextToViewModelLookup.ContainsKey(context))
                return;

            var viewModelInstance = contextToViewModelLookup[context];

            var notifyOfPropertyChanged = viewModelInstance as INotifyPropertyChanged;
            if (notifyOfPropertyChanged != null)
                notifyOfPropertyChanged.PropertyChanged -= NotifiesOfPropertyChangedPropertyChanged;

            var disposible = viewModelInstance as IDisposable;
            if (disposible != null) disposible.Dispose();
            viewModelInstance.Cleanup();

            contextToViewModelLookup.Remove(context);
        }

        public void RegisterCallbackControl(TRibbonTypes ribbonType, string controlCallback, string ribbonControl)
        {
            var type = ribbonTypeLookup[ribbonType];
            if (!notifyChangeTargetLookup.ContainsKey(type))
                notifyChangeTargetLookup.Add(type, new List<KeyValuePair<string, string>>());

            notifyChangeTargetLookup[type].Add(new KeyValuePair<string, string>(controlCallback, ribbonControl));
        }
    }
}
