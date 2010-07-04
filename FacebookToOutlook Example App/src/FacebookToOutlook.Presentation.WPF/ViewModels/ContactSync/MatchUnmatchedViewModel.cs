using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using Bindable.Linq;
using Bindable.Linq.Interfaces;
using FacebookToOutlook.Core;
using FacebookToOutlook.Presentation.Commands;
using FacebookToOutlook.Services;
using GalaSoft.MvvmLight;

namespace FacebookToOutlook.Presentation.ViewModels.ContactSync
{
    public class MatchUnmatchedViewModel : ViewModelBase
    {
        private readonly ObservableCollection<IFacebookUser> _unmatchedContacts;
        private readonly IFacebookUser _userToMatch;
        private readonly IDialogService _dialogService;
        private readonly Action<IFacebookUser> _createNewUser;
        private DelegateCommand _selectContactCommand;
        private DelegateCommand _createNewContactCommand;
        private IFacebookUser _selectedContact;
        private string _searchText = string.Empty;

        public MatchUnmatchedViewModel(IEnumerable<IFacebookUser> unmatchedContacts, IFacebookUser userToMatch,
            IDialogService dialogService) : this(unmatchedContacts, userToMatch, dialogService, null)
        { }

        public MatchUnmatchedViewModel(IEnumerable<IFacebookUser> unmatchedContacts, IFacebookUser userToMatch, 
            IDialogService dialogService, Action<IFacebookUser> createNewUser)
        {
            _unmatchedContacts = new ObservableCollection<IFacebookUser>(unmatchedContacts.OrderBy(c => c.Name));
            _userToMatch = userToMatch;
            _dialogService = dialogService;
            _createNewUser = createNewUser;
        }

        public string SearchText
        {
            get { return _searchText; }
            set
            {
                _searchText = value;
                RaisePropertyChanged("SearchText");
            }
        }

        public IBindableCollection<IFacebookUser> UnmatchedContacts
        {
            get { return _unmatchedContacts.AsBindable().Where(c =>
                (c.Name != null && c.Name.IndexOf(SearchText, StringComparison.CurrentCultureIgnoreCase) != -1) ||
                (c.Company != null && c.Company.IndexOf(SearchText, StringComparison.CurrentCultureIgnoreCase) != -1));
            }
        }

        public IFacebookUser UserToMatch
        {
            get { return _userToMatch; }
        }

        public ICommand SelectContactCommand
        {
            get
            {
                return _selectContactCommand ??
                    (_selectContactCommand = new DelegateCommand(() => _dialogService.CloseDialog(this, true), ()=>SelectedContact != null));
            }
        }

        public ICommand CreateNewContactCommand
        {
            get
            {
                return _createNewContactCommand ??
                    (_createNewContactCommand = new DelegateCommand(()=>_createNewUser(_userToMatch)));
            }
        }

        public bool CanCreateNewContact
        {
            get { return _createNewUser != null; }
        }

        public IFacebookUser SelectedContact
        {
            get { return _selectedContact; }
            set
            {
                _selectedContact = value;
                RaisePropertyChanged("SelectedContact");
            }
        }
    }
}
