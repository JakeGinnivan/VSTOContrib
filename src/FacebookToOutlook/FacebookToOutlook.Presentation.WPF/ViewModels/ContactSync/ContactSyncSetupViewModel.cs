using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using Bindable.Linq;
using Bindable.Linq.Interfaces;
using FacebookToOutlook.Core;
using FacebookToOutlook.Presentation.Commands;
using FacebookToOutlook.Services;
using GalaSoft.MvvmLight;

namespace FacebookToOutlook.Presentation.ViewModels.ContactSync
{
    public class ContactSyncSetupViewModel : ViewModelBase
    {
        private readonly Dispatcher _uiDispatcher;
        private readonly ContactListsBuilder _contactListBuilder;
        private readonly Presentation.ContactSync _contactSync;
        private readonly IDialogService _dialogService;
        private DelegateCommand _matchUnmatchedFriendCommand;
        private DelegateCommand _syncContactsCommand;
        private DelegateCommand _cancelSyncCommand;
        private DelegateCommand _switchListsCommand;
        private bool _unmatchedListEnabled = true;
        private IFacebookUser _selectedUnmatchedContact;
        private readonly IBindable<bool?> _selectionState;
        private bool _isLoading = true;
        private readonly ObservableCollection<MatchedUserViewModel> _matchedContacts = new ObservableCollection<MatchedUserViewModel>();
        private UnmatchedContactsViewModel _unmatchedUsers;

        public ContactSyncSetupViewModel(Dispatcher uiDispatcher, ContactListsBuilder contactListBuilder, 
            Presentation.ContactSync contactSync,
            IDialogService dialogService)
        {
            _uiDispatcher = uiDispatcher;
            _contactListBuilder = contactListBuilder;
            _contactSync = contactSync;
            _dialogService = dialogService;

            _contactListBuilder.BuildComplete += ContactListBuilderBuildComplete;
            _contactListBuilder.Build();

            _selectionState = _matchedContacts
                    .AsBindable()
                    .Count(c => c.SynchroniseMatch)
                    .Switch()
                    .Case<int, bool?>(0, false)
                    .Case(c=>c == MatchedContacts.Count, true)
                    .Default((bool?)null)
                    .EndSwitch();
            _selectionState.PropertyChanged += SelectionStatePropertyChanged;
        }

        void ContactListBuilderBuildComplete(object sender, BuildCompleteEventArgs e)
        {
            if (!e.Success)
            {
                MessageBox.Show(e.ErrorMessage, "Error loading contacts");
                _uiDispatcher.Invoke(((Action)(() => _dialogService.Close(this))));
                return;
            }

            foreach (var matchedUserViewModel in e.MatchedUsers)
                _matchedContacts.Add(matchedUserViewModel);

            UnmatchedUsers = e.UnmatchedList;
            UnmatchedUsers.NewMatch += UnmatchedUsers_NewMatch;
            IsLoading = false;
            RaisePropertyChanged("UnmatchedListEnabled");
        }

        void UnmatchedUsers_NewMatch(object sender, NewMatchEventArgs e)
        {
            _matchedContacts.Add(e.NewMatch);
        }

        void SelectionStatePropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            RaisePropertyChanged("SelectionState");
        }

        public bool IsLoading
        {
            get { return _isLoading; }
            private set
            {
                _isLoading = value;
                RaisePropertyChanged("IsLoading");
            }
        }

        public ICommand SwitchListsCommand
        {
            get
            {
                return _switchListsCommand ??
            
                (_switchListsCommand = new DelegateCommand(()=>
                                                               {
                                                                   SelectedUnmatchedContact = null;
                                                                   _unmatchedUsers.SwitchLists();
                                                               }));
            }
        }

        

        public ICommand SyncContactsCommand
        {
            get
            {
                return _syncContactsCommand ??
                    (_syncContactsCommand = new DelegateCommand(SyncContacts, () => !IsLoading && SelectionState != false));
            }
        }

        private void SyncContacts()
        {
            var users =
                _matchedContacts
                .Where(c => c.SynchroniseMatch);

            _contactSync.Sync(users);
            _dialogService.Close(this);
        }

        public ICommand CancelSyncCommand
        {
            get
            {
                return _cancelSyncCommand ??
                       (_cancelSyncCommand =
                        new DelegateCommand(() => _dialogService.CloseDialog(this, false), () => !IsLoading));
            }
        }

        public ICommand MatchUnmatchedFriendCommand
        {
            get
            {
                return _matchUnmatchedFriendCommand ??
                       (_matchUnmatchedFriendCommand = new DelegateCommand(MatchUnmatchedFriend, () => !IsLoading));
            }
        }

        private void MatchUnmatchedFriend()
        {
            UnmatchedListEnabled = false;

            try
            {
                _unmatchedUsers.MatchUnmatchedFriend(SelectedUnmatchedContact);
            }
            finally
            {
                UnmatchedListEnabled = true;
            }
        }

        public IFacebookUser SelectedUnmatchedContact
        {
            get {
                return _selectedUnmatchedContact;
            }
            set {
                _selectedUnmatchedContact = value;
                RaisePropertyChanged("SelectedUnmatchedContact");
            }
        }

        public bool UnmatchedListEnabled
        {
            get
            {
                return !IsLoading && _unmatchedListEnabled;
            }
            private set 
            {
                _unmatchedListEnabled = value;
                RaisePropertyChanged("UnmatchedListEnabled");
            }
        }
        
        public IBindableCollection<MatchedUserViewModel> MatchedContacts
        {
            get
            {
                return _matchedContacts.AsBindable().OrderBy(c => c.OutlookContact.Name);
            }
        }

        public bool? SelectionState
        {
            get
            {
                return _selectionState.Current;
            }
            set
            {
                _selectionState.PropertyChanged -= SelectionStatePropertyChanged;

                if (value != null)
                {
                    foreach (var matchedOutlookContact in MatchedContacts)
                    {
                        matchedOutlookContact.SynchroniseMatch = value.Value;
                    }
                }

                _selectionState.PropertyChanged += SelectionStatePropertyChanged;
                
            }
        }

        public UnmatchedContactsViewModel UnmatchedUsers
        {
            get { return _unmatchedUsers; }
            set
            {
                _unmatchedUsers = value;
                RaisePropertyChanged("UnmatchedUsers");
            }
        }
    }
}
