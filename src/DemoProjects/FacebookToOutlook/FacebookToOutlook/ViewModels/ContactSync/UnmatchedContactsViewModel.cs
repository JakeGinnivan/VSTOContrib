using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Linq.Expressions;
using System.Windows;
using Bindable.Linq;
using Bindable.Linq.Interfaces;
using FacebookToOutlook.Core;
using FacebookToOutlook.Data;
using FacebookToOutlook.Services;
using FacebookToOutlook.Views;
using GalaSoft.MvvmLight;

namespace FacebookToOutlook.Presentation.ViewModels.ContactSync
{
    public class UnmatchedContactsViewModel : ViewModelBase
    {
        private readonly IDialogService _dialogService;
        private readonly IOutlookRepository _outlookRepository;
        private readonly ObservableCollection<IFacebookUser> _unmatchedFacebookUsers;
        private readonly ObservableCollection<IOutlookFacebookUser> _unmatchedOutlookContacts;
        private CurrentList _currentUnmatchedList = CurrentList.FacebookUsers;
        public event EventHandler<NewMatchEventArgs> NewMatch;
        private string _searchText = string.Empty;

        public UnmatchedContactsViewModel(
            IDialogService dialogService, 
            IOutlookRepository outlookRepository,
            ObservableCollection<IFacebookUser> unmatchedFacebookUsers, 
            ObservableCollection<IOutlookFacebookUser> unmatchedOutlookContacts)
        {
            _dialogService = dialogService;
            _outlookRepository = outlookRepository;
            _unmatchedFacebookUsers = unmatchedFacebookUsers;
            _unmatchedOutlookContacts = unmatchedOutlookContacts;
        }

        public CurrentList CurrentUnmatchedList
        {
            get { return _currentUnmatchedList; }
            set
            {
                _currentUnmatchedList = value;
                RaisePropertyChanged("CurrentUnmatchedList");
                RaisePropertyChanged("UnmatchedText");
                RaisePropertyChanged("SwitchListsText");
                RaisePropertyChanged("CurrentUnmatchedListContacts");
            }
        }

        public string UnmatchedText
        {
            get
            {
                return CurrentUnmatchedList == CurrentList.FacebookUsers ? "Unmatched facebook users" : "Unmatched outlook contacts";
            }
        }

        public string SwitchListsText
        {
            get
            {
                return CurrentUnmatchedList == CurrentList.FacebookUsers ? "Switch to Outlook Contacts" : "Switch to Facebook Friends";
            }
        }

        public void SwitchLists()
        {
            CurrentUnmatchedList = (CurrentUnmatchedList == CurrentList.FacebookUsers
                                        ? CurrentList.OutlookContacts
                                        : CurrentList.FacebookUsers);
        }

        public void MatchUnmatchedFriend(IFacebookUser userToMatch)
        {
            var matchUnmatchedViewModel = GetMatchUnmatchedViewModel(userToMatch);

            if (_dialogService.ShowDialog<MatchUnmatchedView>(null, matchUnmatchedViewModel) == true)
            {
                MatchUsers(userToMatch, matchUnmatchedViewModel.SelectedContact);
            }
        }

        private void MatchUsers(IFacebookUser userToMatch, IFacebookUser matchTo)
        {
            IFacebookUser facebookUser;
            IOutlookFacebookUser outlookFacebookUser;

            //Assign correct user to variables for association
            if (CurrentUnmatchedList == CurrentList.FacebookUsers)
            {
                outlookFacebookUser = (IOutlookFacebookUser)matchTo;
                facebookUser = userToMatch;
            }
            else
            {
                outlookFacebookUser = (IOutlookFacebookUser)userToMatch;
                facebookUser = matchTo;
            }

            OfferToCorrectName(facebookUser, outlookFacebookUser);

            _outlookRepository.AssociateFacebookUserWithContact(outlookFacebookUser, facebookUser);
            _unmatchedOutlookContacts.Remove(outlookFacebookUser);
            _unmatchedFacebookUsers.Remove(facebookUser);
            RaiseNewUserMatch(new MatchedUserViewModel(outlookFacebookUser, facebookUser, true));
        }

        public string SearchText
        {
            get
            {
                return _searchText;
            }
            set
            {
                _searchText = value;
                RaisePropertyChanged("SearchText");
            }
        }

        public IBindableCollection<IFacebookUser> CurrentUnmatchedListContacts
        {
            get
            {
                return CurrentUnmatchedList == CurrentList.FacebookUsers ? UnmatchedFacebookContacts : UnmatchedOutlookContacts;
            }
        }

        public IBindableCollection<IFacebookUser> UnmatchedOutlookContacts
        {
            get
            {
                return
                    _unmatchedOutlookContacts
                        .AsBindable()
                        .Select(c => (IFacebookUser)c)
                        .Where(SearchExpression())
                        .OrderBy(c => c.Name);
            }
        }

        private Expression<Func<IFacebookUser, bool>> SearchExpression()
        {
            return c =>
                   (c.Name != null && c.Name.IndexOf(SearchText, StringComparison.CurrentCultureIgnoreCase) != -1) ||
                   (c.Company != null && c.Company.IndexOf(SearchText, StringComparison.CurrentCultureIgnoreCase) != -1);
        }

        public IBindableCollection<IFacebookUser> UnmatchedFacebookContacts
        {
            get
            {
                return
                    _unmatchedFacebookUsers
                        .AsBindable()
                        .Select(c => c)
                        .Where(SearchExpression())
                        .OrderBy(c => c.Name);
            }
        }

        private MatchUnmatchedViewModel GetMatchUnmatchedViewModel(IFacebookUser userToMatch)
        {
            MatchUnmatchedViewModel matchUnmatchedViewModel = null;
            if (CurrentUnmatchedList == CurrentList.FacebookUsers)
            {
                matchUnmatchedViewModel = new MatchUnmatchedViewModel(
                    _unmatchedOutlookContacts.Cast<IFacebookUser>(), userToMatch,
                    _dialogService,
                    //createNewUser Action:
                    u =>
                    {
                        // ReSharper disable AccessToModifiedClosure
                        _dialogService.CloseDialog(matchUnmatchedViewModel, null);
                        // ReSharper restore AccessToModifiedClosure
                        var outlookContact = _outlookRepository.CreateContactFromFacebookUser(u);
                        _unmatchedFacebookUsers.Remove(userToMatch);
                        RaiseNewUserMatch(new MatchedUserViewModel(outlookContact, userToMatch, true));
                    });
            }
            else
            {
                matchUnmatchedViewModel = new MatchUnmatchedViewModel(_unmatchedFacebookUsers, userToMatch, _dialogService);
            }
            return matchUnmatchedViewModel;
        }

        private void OfferToCorrectName(IFacebookUser facebookUser, IFacebookUser outlookFacebookUser)
        {
            if (facebookUser.Name == outlookFacebookUser.Name) return;

            var messageBoxText = string.Format(Properties.Resources.NameMismatchFix, facebookUser.Name, outlookFacebookUser.Name);
            var messageBoxResult = _dialogService.ShowMessageBox(this, messageBoxText, "Update name in outlook", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (messageBoxResult == MessageBoxResult.Yes)
            {
                outlookFacebookUser.Name = facebookUser.Name;
            }
        }

        private void RaiseNewUserMatch(MatchedUserViewModel matchedUserViewModel)
        {
            if (NewMatch == null)
                throw new InvalidOperationException("The NewUser event must be handled");

            NewMatch(this, new NewMatchEventArgs(matchedUserViewModel));
        }
    }
}
