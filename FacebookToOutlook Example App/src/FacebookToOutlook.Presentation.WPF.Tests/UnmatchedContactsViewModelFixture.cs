using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using FacebookToOutlook.Core;
using FacebookToOutlook.Data;
using FacebookToOutlook.Presentation.ViewModels.ContactSync;
using FacebookToOutlook.Presentation.Views;
using FacebookToOutlook.Services;
using Xunit;
using Rhino.Mocks;

namespace FacebookToOutlook.Presentation.WPF.Tests
{
    public class UnmatchedContactsViewModelFixture
    {
        [Fact]
        public void CurrentUnmatchedListIsFacebookListByDefault()
        {
            var unmatchedContacts = new UnmatchedContactsViewModel(null, null, null, null);
            Assert.Equal(CurrentList.FacebookUsers, unmatchedContacts.CurrentUnmatchedList);
        }

        [Fact]
        public void SwitchListsChangesFromFacebookUsersToOutlookUsers()
        {
            var unmatchedContacts = new UnmatchedContactsViewModel(null, null, null, null);

            unmatchedContacts.SwitchLists();

            Assert.Equal(CurrentList.OutlookContacts, unmatchedContacts.CurrentUnmatchedList);
        }

        [Fact]
        public void TwoListSwitchesIsFacebookUsersList()
        {
            var unmatchedContacts = new UnmatchedContactsViewModel(null, null, null, null);

            unmatchedContacts.SwitchLists();
            unmatchedContacts.SwitchLists();

            Assert.Equal(CurrentList.FacebookUsers, unmatchedContacts.CurrentUnmatchedList);
        }

        [Fact]
        public void CurrentUnmatchedListContactsReturnsFacebookListWhenCurrent()
        {
            var unmatchedContacts = new UnmatchedContactsViewModel(null, null, new ObservableCollection<IFacebookUser> { new FacebookUser { Name = "Jake" } }, null);

            Assert.Equal("Jake", unmatchedContacts.CurrentUnmatchedListContacts.First().Name);
        }

        [Fact]
        public void CurrentUnmatchedListContactsReturnsOutlookListWhenCurrent()
        {
            var unmatchedContacts = new UnmatchedContactsViewModel(null, null, null, new ObservableCollection<IOutlookFacebookUser>
                                                                                         {
                                                                                             new OutlookFacebookUser(string.Empty) { Name = "Jake" }
                                                                                         });

            unmatchedContacts.SwitchLists();

            Assert.Equal("Jake", unmatchedContacts.CurrentUnmatchedListContacts.First().Name);
        }

        [Fact]
        public void CurrentUnmatchedListRaisesPropertyEventChangedForCurrentUnmatchedList()
        {
            var unmatchedContacts = new UnmatchedContactsViewModel(null, null, null, null);
            AssertUtil.PropertyChangedEvent(unmatchedContacts, c => c.CurrentUnmatchedList, () => unmatchedContacts.CurrentUnmatchedList = CurrentList.FacebookUsers);
        }

        [Fact]
        public void CurrentUnmatchedListRaisesPropertyEventChangedForUnmatchedText()
        {
            var unmatchedContacts = new UnmatchedContactsViewModel(null, null, null, null);
            AssertUtil.PropertyChangedEvent(unmatchedContacts, c => c.UnmatchedText, () => unmatchedContacts.CurrentUnmatchedList = CurrentList.FacebookUsers);
        }

        [Fact]
        public void CurrentUnmatchedListRaisesPropertyEventChangedForSwitchListsText()
        {
            var unmatchedContacts = new UnmatchedContactsViewModel(null, null, null, null);
            AssertUtil.PropertyChangedEvent(unmatchedContacts, c => c.SwitchListsText, () => unmatchedContacts.CurrentUnmatchedList = CurrentList.FacebookUsers);
        }

        [Fact]
        public void CurrentUnmatchedListRaisesPropertyEventChangedForCurrentUnmatchedListContacts()
        {
            var unmatchedContacts = new UnmatchedContactsViewModel(null, null, null, null);
            AssertUtil.PropertyChangedEvent(unmatchedContacts, c => c.CurrentUnmatchedListContacts, () => unmatchedContacts.CurrentUnmatchedList = CurrentList.FacebookUsers);
        }

        [Fact]
        public void SearchTextFiltersByName()
        {
            var unmatchedContacts = new UnmatchedContactsViewModel(null, null, new ObservableCollection<IFacebookUser>
                                                                                   {
                                                                                       new FacebookUser { Name = "Jake" },
                                                                                       new FacebookUser{ Name = "Bob"}
                                                                                   }, null){SearchText = "ja"};

            Assert.Equal("Jake", unmatchedContacts.CurrentUnmatchedListContacts.Single().Name);
        }

        [Fact]
        public void MatchUnmatchedUserCallsNewMatchEvent()
        {
            //Arrange
            var facebookUser = new FacebookUser {Name = "Jake G", UserId = 1};
            var outlookFacebookUser = new OutlookFacebookUser(string.Empty){Name = "Jake"};
            var unmatchedFacebookUsers = new ObservableCollection<IFacebookUser>{facebookUser};
            var unmatchedOutlookContacts = new ObservableCollection<IOutlookFacebookUser>{outlookFacebookUser};

            var dialogService = MockRepository.GenerateStub<IDialogService>();
            dialogService
                .Stub(d => d.ShowDialog<MatchUnmatchedView>(Arg<object>.Is.Anything, Arg<object>.Is.Anything))
                .Do((Func<object, object, bool?>)((parent, viewModel) =>
                                                                       {
                                                                           ((MatchUnmatchedViewModel)viewModel).SelectedContact = outlookFacebookUser;
                                                                           return true;
                                                                       }));
            var outlookRepository = MockRepository.GenerateStub<IOutlookRepository>();
            var unmatchedContacts = new UnmatchedContactsViewModel(dialogService, outlookRepository,
                                                                   unmatchedFacebookUsers,
                                                                   unmatchedOutlookContacts);

            var newMatchCalled = false;
            unmatchedContacts.NewMatch += (sender, e) => { newMatchCalled = true; };

            //Act
            unmatchedContacts.MatchUnmatchedFriend(facebookUser);

            //Assert
            Assert.True(newMatchCalled);
        }

        [Fact]
        public void MatchUnmatchedUserAssociatesUserInOutlookRepository()
        {
            //Arrange
            var facebookUser = new FacebookUser {Name = "Jake G", UserId = 1};
            var outlookFacebookUser = new OutlookFacebookUser(string.Empty) {Name = "Jake"};
            var unmatchedFacebookUsers = new ObservableCollection<IFacebookUser> {facebookUser};
            var unmatchedOutlookContacts = new ObservableCollection<IOutlookFacebookUser> {outlookFacebookUser};

            var dialogService = MockRepository.GenerateStub<IDialogService>();
            dialogService
                .Stub(d => d.ShowDialog<MatchUnmatchedView>(Arg<object>.Is.Anything, Arg<object>.Is.Anything))
                .Do((Func<object, object, bool?>) ((parent, viewModel) =>
                                                       {
                                                           ((MatchUnmatchedViewModel) viewModel).SelectedContact =
                                                               outlookFacebookUser;
                                                           return true;
                                                       }));
            var outlookRepository = MockRepository.GenerateMock<IOutlookRepository>();

            var unmatchedContacts = new UnmatchedContactsViewModel(dialogService, outlookRepository, unmatchedFacebookUsers, unmatchedOutlookContacts);

            unmatchedContacts.NewMatch += (sender, e) => { };

            //Act
            unmatchedContacts.MatchUnmatchedFriend(facebookUser);

            outlookRepository.AssertWasCalled(r => r.AssociateFacebookUserWithContact(outlookFacebookUser, facebookUser));
        }

        [Fact]
        public void MatchUnmatchedUserOffersToCorrectName()
        {
            //Arrange
            var facebookUser = new FacebookUser { Name = "Jake G", UserId = 1 };
            var outlookFacebookUser = new OutlookFacebookUser(string.Empty) { Name = "Jake" };
            var unmatchedFacebookUsers = new ObservableCollection<IFacebookUser> { facebookUser };
            var unmatchedOutlookContacts = new ObservableCollection<IOutlookFacebookUser> { outlookFacebookUser };

            var dialogService = MockRepository.GenerateMock<IDialogService>();
            dialogService
                .Stub(d => d.ShowDialog<MatchUnmatchedView>(Arg<object>.Is.Anything, Arg<object>.Is.Anything))
                .Do((Func<object, object, bool?>)((parent, viewModel) =>
                {
                    ((MatchUnmatchedViewModel)viewModel).SelectedContact = outlookFacebookUser;
                    return true;
                }));
            var outlookRepository = MockRepository.GenerateStub<IOutlookRepository>();

            var unmatchedContacts = new UnmatchedContactsViewModel(dialogService, outlookRepository, unmatchedFacebookUsers, unmatchedOutlookContacts);

            unmatchedContacts.NewMatch += (sender, e) => { };

            //Act
            unmatchedContacts.MatchUnmatchedFriend(facebookUser);

            dialogService.AssertWasCalled(d => d.ShowMessageBox(Arg<object>.Is.Anything, Arg<string>.Is.Anything, Arg<string>.Is.Anything
                , Arg<MessageBoxButton>.Is.Equal(MessageBoxButton.YesNo), Arg<MessageBoxImage>.Is.Anything));
        }
    }
}
