﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Microsoft.Office.Core;
using VSTOContrib.Core.Annotations;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Interfaces.Internal;

namespace VSTOContrib.Core.RibbonFactory.Internal
{
    internal class ViewModelResolver : IViewModelResolver
    {
        /// <summary>
        /// Used when a new explorer or inspector is created to lookup the appropriate viewmodel type
        /// </summary>
        readonly Dictionary<string, Type> ribbonTypeLookup;
        /// <summary>
        /// Internal lookup for Context instances to view model lookups
        /// </summary>
        readonly Dictionary<object, IRibbonViewModel> contextToViewModelLookup;
        readonly Dictionary<string, IRibbonUI> ribbonUiLookup;
        /// <summary>
        /// Looks up ViewModelType, callback method name, control id, controlId used to invalidate :)
        /// </summary>
        readonly Dictionary<Type, List<KeyValuePair<string,string>>> notifyChangeTargetLookup;

        readonly ICustomTaskPaneRegister customTaskPaneRegister;
        readonly IViewContextProvider viewContextProvider;
        readonly VstoContribContext vstoContribContext;
        IViewProvider viewProvider;
        string currentlyLoadingRibbon;

        public ViewModelResolver(IEnumerable<Type> viewModelType, ICustomTaskPaneRegister customTaskPaneRegister, IViewContextProvider viewContextProvider, VstoContribContext context)
        {
            currentlyLoadingRibbon = "default";
            notifyChangeTargetLookup = new Dictionary<Type, List<KeyValuePair<string, string>>>();
            ribbonTypeLookup = new Dictionary<string, Type>();
            contextToViewModelLookup = new Dictionary<object, IRibbonViewModel>();
            ribbonUiLookup = new Dictionary<string, IRibbonUI>();
            this.customTaskPaneRegister = customTaskPaneRegister;
            this.viewContextProvider = viewContextProvider;
            vstoContribContext = context;

            foreach (var ribbonType in viewModelType)
            {
                CreateRibbonTypeToViewModelTypeLookup(ribbonType, context.FallbackRibbonType);
            }
        }

        public void Initialise(IViewProvider viewProvider)
        {
            this.viewProvider = viewProvider;

            this.viewProvider.NewView += ViewProviderNewView;
            this.viewProvider.ViewClosed += ViewProviderViewClosed;
            this.viewProvider.UpdateCustomTaskPanesVisibilityForContext += 
                ViewProviderOnUpdateCustomTaskPanesVisibilityForContext;
        }

        void ViewProviderOnUpdateCustomTaskPanesVisibilityForContext(object sender, 
            HideCustomTaskPanesForContextEventArgs hideCustomTaskPanesForContextEventArgs)
        {
            customTaskPaneRegister.ChangeVisibilityForContext(hideCustomTaskPanesForContextEventArgs.Context,
                hideCustomTaskPanesForContextEventArgs.Visible);
        }

        void ViewProviderViewClosed(object sender, ViewClosedEventArgs e)
        {
            //TODO write test around context/view cleanup
            customTaskPaneRegister.Cleanup(e.View);

            CleanupViewModel(e.Context);
            viewProvider.CleanupReferencesTo(e.View, e.Context);
        }

        void ViewProviderNewView(object sender, NewViewEventArgs e)
        {
            var viewModel = GetOrCreateViewModel(e);
            if (viewModel == null) return;
            customTaskPaneRegister.RegisterCustomTaskPanes(viewModel, e.ViewInstance, e.ViewContext);
            e.Handled = true;
        }

        IRibbonViewModel GetOrCreateViewModel(NewViewEventArgs e)
        {
            if (!ribbonTypeLookup.ContainsKey(e.RibbonType)) return null;
            if (contextToViewModelLookup.ContainsKey(e.ViewContext))
            {
                //Tell viewmodel there is a new view active
                var ribbonViewModel = contextToViewModelLookup[e.ViewContext];
                ribbonViewModel.CurrentViewChanged(e.ViewInstance);
                return ribbonViewModel;
            }

            currentlyLoadingRibbon = e.RibbonType;
            IRibbonViewModel buildViewModel = BuildViewModel(e.RibbonType, e.ViewInstance, e.ViewContext);
            contextToViewModelLookup.Add(e.ViewContext, buildViewModel);
            return buildViewModel;
        }

        private void CreateRibbonTypeToViewModelTypeLookup(Type ribbonViewModel, [CanBeNull] string ribbonFallbackType)
        {
            foreach (var value in ViewModelRibbonTypesLookupProvider.Instance.GetRibbonTypesFor(ribbonViewModel, ribbonFallbackType))
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
                var ribbonTypeForView = viewContextProvider.GetRibbonTypeForView(view);
                var newViewEventArgs = new NewViewEventArgs(view, context, ribbonTypeForView);

                GetOrCreateViewModel(newViewEventArgs);
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

        private IRibbonViewModel BuildViewModel(string ribbonType, object viewInstance, object viewContext)
        {
            var viewModelType = ribbonTypeLookup[ribbonType];
            var ribbonViewModel = vstoContribContext.ViewModelFactory.Resolve(viewModelType);
            ribbonViewModel.VstoFactory = vstoContribContext.VstoFactory;

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

            viewModelInstance.Cleanup();
            vstoContribContext.ViewModelFactory.Release(viewModelInstance);

            contextToViewModelLookup.Remove(context);
        }

        public void RegisterCallbackControl(string ribbonType, string controlCallback, string ribbonControl)
        {
            var type = ribbonTypeLookup[ribbonType];
            if (!notifyChangeTargetLookup.ContainsKey(type))
                notifyChangeTargetLookup.Add(type, new List<KeyValuePair<string, string>>());

            notifyChangeTargetLookup[type].Add(new KeyValuePair<string, string>(controlCallback, ribbonControl));
        }

        public void Dispose()
        {
            var disposable = vstoContribContext.ViewModelFactory as IDisposable;
            if (disposable != null)
                disposable.Dispose();
        }
    }
}
