using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading;
using FacebookToOutlook.Core;
using FacebookToOutlook.Data;
using FacebookToOutlook.Presentation.ViewModels.ContactSync;
using FacebookToOutlook.Services;
using Rhino.Mocks;
using Xunit;

namespace FacebookToOutlook.Presentation.WPF.Tests
{
    public class ContactListsBuilderFixture
    {
        private readonly IOutlookRepository _outlookRepository;
        private readonly IFacebookRepository _facebookRepository;
        private readonly Func<ObservableCollection<IFacebookUser>, ObservableCollection<IOutlookFacebookUser>, UnmatchedContactsViewModel> _unmatchedListFactory;

        public ContactListsBuilderFixture()
        {
            _outlookRepository = MockRepository.GenerateMock<IOutlookRepository>();
            _facebookRepository = MockRepository.GenerateMock<IFacebookRepository>();
            _unmatchedListFactory = (f, o) => new UnmatchedContactsViewModel(MockRepository.GenerateMock<IDialogService>(), _outlookRepository, f, o);
        }

        public BuildCompleteEventArgs GetBuildResults()
        {
            using (var contactListBuilder = new ContactListsBuilder(_outlookRepository, _facebookRepository, _unmatchedListFactory))
            using (var eventWaitHandle = new ManualResetEventSlim())
            {
                BuildCompleteEventArgs results = null;
                EventHandler<BuildCompleteEventArgs> contactListBuilderOnBuildComplete = (sender, e) =>
                                                                     {
                                                                         results = e;
                                                                         eventWaitHandle.Set();
                                                                     };
                contactListBuilder.BuildComplete += contactListBuilderOnBuildComplete;
                contactListBuilder.Build();
                Assert.True(eventWaitHandle.Wait(10000), "BuildComplete event not fired");
                return results;
            }
        }

        [Fact]
        public void ContactListBuilderMatchesOnName()
        {

            _outlookRepository.Stub(r => r.GetContacts()).Return(new List<IOutlookFacebookUser>
                                                                     {
                                                                         new OutlookFacebookUser(string.Empty)
                                                                             {Name = "Test Name"}
                                                                     });

            _facebookRepository.Stub(r => r.GetFriends()).Return(new List<IFacebookUser>
                                                                     {
                                                                         new FacebookUser {Name = "Test Name", UserId = 1}
                                                                     });

            var buildResults = GetBuildResults();

            Assert.Equal(1, buildResults.MatchedUsers.Count);
        }

        [Fact]
        public void ContactListBuilderMatchesExistingMatchedContact()
        {

            _outlookRepository.Stub(r => r.GetContacts()).Return(new List<IOutlookFacebookUser>
                                                                     {
                                                                         new OutlookFacebookUser(string.Empty)
                                                                             {Name = "Test Name", UserId = 1}
                                                                     });

            _facebookRepository.Stub(r => r.GetFriends()).Return(new List<IFacebookUser>
                                                                     {
                                                                         new FacebookUser {Name = "Different Name", UserId = 1}
                                                                     });

            var buildResults = GetBuildResults();

            Assert.Equal(1, buildResults.MatchedUsers.Count);
        }

        [Fact]
        public void ContactListAddsUnmatchedFacebookContactToList()
        {

            _outlookRepository.Stub(r => r.GetContacts()).Return(new List<IOutlookFacebookUser>());

            _facebookRepository.Stub(r => r.GetFriends()).Return(new List<IFacebookUser>
                                                                     {
                                                                         new FacebookUser {Name = "Name", UserId = 1}
                                                                     });

            var buildResults = GetBuildResults();

            Assert.Equal(1, buildResults.UnmatchedList.UnmatchedFacebookContacts.Count);
        }

        [Fact]
        public void ContactListAddsUnmatchedOutlookContactToList()
        {

            _outlookRepository.Stub(r => r.GetContacts()).Return(new List<IOutlookFacebookUser>
                                                                     {
                                                                         new OutlookFacebookUser(string.Empty) {Name = "Name"}
                                                                     });

            _facebookRepository.Stub(r => r.GetFriends()).Return(new List<IFacebookUser>());

            var buildResults = GetBuildResults();

            Assert.Equal(1, buildResults.UnmatchedList.UnmatchedOutlookContacts.Count);
        }

        [Fact]
        public void ContactListBuilderFetchesListsConcurrently()
        {
            using (var contactListBuilder = new ContactListsBuilder(_outlookRepository, _facebookRepository, _unmatchedListFactory))
            {
                var eventWaitHandle = new ManualResetEventSlim();
                var outlookRepoWaitHandle = new ManualResetEventSlim();
                contactListBuilder.BuildComplete += (sender, e) => eventWaitHandle.Set();
                var outlookThreadId = -1;
                var facebookTheadId = -1;
                _outlookRepository.Stub(r => r.GetContacts()).Do((Func<IList<IOutlookFacebookUser>>)(() =>
                {
                    outlookThreadId = Thread.CurrentThread.ManagedThreadId;
                    //Block until facebook repo unblocks. Stops Race condition where this thread is released
                    //back into threadpool before BeginInvoke call on facebook repo gets thread
                    outlookRepoWaitHandle.Wait();
                    return new List<IOutlookFacebookUser>();
                }));
                _facebookRepository.Stub(r => r.GetFriends()).Do((Func<IList<IFacebookUser>>)(() =>
                {
                    facebookTheadId = Thread.CurrentThread.ManagedThreadId;
                    outlookRepoWaitHandle.Set();
                    return new List<IFacebookUser>();
                }));

                contactListBuilder.Build();
                eventWaitHandle.Wait();

                Assert.NotEqual(outlookThreadId, facebookTheadId);
            }
        }
    }
}
