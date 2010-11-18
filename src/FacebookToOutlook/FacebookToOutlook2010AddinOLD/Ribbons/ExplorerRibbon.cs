using System;
using FacebookToOutlook;
using FacebookToOutlook.Presentation.ViewModels;
using FacebookToOutlook.Presentation.ViewModels.ContactSync;
using FacebookToOutlook.Services;
using FacebookToOutlook.Views;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Office.Utility;
using Outlook.Utility.Services;

namespace FacebookToOutlookAddin.Ribbons
{
    public class ExplorerRibbon : IRibbonViewModel
    {
        private readonly IDialogService _dialogService;
        private readonly FacebookEventSynchronisationService _facebookSyncService;
        private readonly Func<ConfigurationViewModel> _configViewModelFactory;
        private readonly Func<ContactSyncSetupViewModel> _syncSetupViewModelFactory;

        public ExplorerRibbon(
            IDialogService dialogService, 
            FacebookEventSynchronisationService facebookSyncService,
            Func<ConfigurationViewModel> configViewModelFactory,
            Func<ContactSyncSetupViewModel> syncSetupViewModelFactory)
        {
            _dialogService = dialogService;
            _facebookSyncService = facebookSyncService;
            _configViewModelFactory = configViewModelFactory;
            _syncSetupViewModelFactory = syncSetupViewModelFactory;
        }

        public RibbonType Type
        {
            get { return RibbonType.OutlookExplorer; }
        }

        public IRibbonUI RibbonUi
        {
            get;
            set;
        }

        public void SyncEventsButtonClick(IRibbonControl ribbonControl)
        {
            if (_facebookSyncService.SyncInProgress) return;
            _facebookSyncService.SynchronisationComplete += FacebookEventSynchronisationServiceSynchronisationComplete;
            _facebookSyncService.SynchroniseAsync();
            RibbonUi.InvalidateControl("syncEventsButton");
        }

        public bool SyncButtonIsEnabled(IRibbonControl control)
        {
            return !_facebookSyncService.SyncInProgress;
        }

        void FacebookEventSynchronisationServiceSynchronisationComplete(object sender, SynchronisationCompleteEventArgs e)
        {
            _facebookSyncService.SynchronisationComplete -= FacebookEventSynchronisationServiceSynchronisationComplete;
            RibbonUi.InvalidateControl("syncEventsButton");
        }

        public void ShowFacebookEventsClick(IRibbonControl ribbonControl)
        {
            var explorer = (Explorer)ribbonControl.Context;
            explorer.Search("[IsFacebookEvent]:true", OlSearchScope.olSearchScopeCurrentFolder);
        }

        public void ShowLinkedContactsClick(IRibbonControl ribbonControl)
        {
            var explorer = (Explorer)ribbonControl.Context;
            explorer.Search("[IsLinkedToFacebookUser]:true", OlSearchScope.olSearchScopeCurrentFolder);
        }

        public void ContactConfigButtonClick(IRibbonControl ribbonControl)
        {
            ShowConfigDialog(ConfigurationTab.ContactsTab);
        }

        public void EventConfigButtonClick(IRibbonControl ribbonControl)
        {
            ShowConfigDialog(ConfigurationTab.EventsTab);
        }

        private void ShowConfigDialog(ConfigurationTab selectedTab)
        {
            var configurationViewModel = _configViewModelFactory();
            configurationViewModel.SelectedConfigurationTab = selectedTab;
            _dialogService.ShowDialog<ConfigurationView>(null, configurationViewModel);
        }

        public void SyncContactsButtonClick(IRibbonControl ribbonControl)
        {
            _dialogService.Show<ContactSyncView>(null, _syncSetupViewModelFactory);
        }
    }
}
