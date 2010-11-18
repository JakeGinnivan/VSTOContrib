using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using FacebookToOutlook.Core;
using FacebookToOutlook.Data;
using FacebookToOutlook.Presentation.ViewModels.ContactSync;

namespace FacebookToOutlook.Presentation
{
    public class ContactListsBuilder : IDisposable
    {
        private readonly IOutlookRepository _outlookRepository;
        private readonly IFacebookRepository _facebookRepository;
        private readonly Func<ObservableCollection<IFacebookUser>, ObservableCollection<IOutlookFacebookUser>, UnmatchedContactsViewModel> _unmatchedListFactory;
        private BackgroundWorker _contactLoader = new BackgroundWorker();
        public event EventHandler<BuildCompleteEventArgs> BuildComplete;

        private readonly ObservableCollection<MatchedUserViewModel> _matchedOutlookContacts = new ObservableCollection<MatchedUserViewModel>();
        private readonly ObservableCollection<IOutlookFacebookUser> _unmatchedOutlookContacts = new ObservableCollection<IOutlookFacebookUser>();
        private readonly ObservableCollection<IFacebookUser> _unmatchedFacebookContacts = new ObservableCollection<IFacebookUser>();
        private bool _disposed;

        public ContactListsBuilder(IOutlookRepository outlookRepository, IFacebookRepository facebookRepository,
            Func<ObservableCollection<IFacebookUser>, ObservableCollection<IOutlookFacebookUser>, UnmatchedContactsViewModel> unmatchedListFactory)
        {
            _outlookRepository = outlookRepository;
            _facebookRepository = facebookRepository;
            _unmatchedListFactory = unmatchedListFactory;
            _contactLoader.DoWork += ContactLoaderDoWork;
            _contactLoader.RunWorkerCompleted += ContactLoaderRunWorkerCompleted;
        }

        public void Build()
        {
            _contactLoader.RunWorkerAsync();
        }

        void ContactLoaderRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (BuildComplete == null) throw new InvalidOperationException("BuildComplete event must be handled");
            var errorMessage = e.Error == null ? null : e.Error.Message;
            var unmatchedList = _unmatchedListFactory(_unmatchedFacebookContacts, _unmatchedOutlookContacts);
            var buildCompleteEventArgs = new BuildCompleteEventArgs(e.Error == null, errorMessage, _matchedOutlookContacts, unmatchedList);
            BuildComplete(this, buildCompleteEventArgs);
        }

        void ContactLoaderDoWork(object sender, DoWorkEventArgs e)
        {
            IList<IOutlookFacebookUser> outlookContacts;
            Dictionary<long, IFacebookUser> facebookContacts;
            GetContacts(out outlookContacts, out facebookContacts);

            for (var index = 0; index < outlookContacts.Count; index++)
            {
                ProcessContactAtIndex(ref index, outlookContacts, facebookContacts);
            }

            foreach (var outlookContact in outlookContacts)
                _unmatchedOutlookContacts.Add(outlookContact);
            foreach (var facebookContact in facebookContacts)
                _unmatchedFacebookContacts.Add(facebookContact.Value);
        }

        private void ProcessContactAtIndex(ref int index, IList<IOutlookFacebookUser> outlookContacts, Dictionary<long, IFacebookUser> facebookContacts)
        {
            var outlookContact = outlookContacts[index];
            var facebookUserId = outlookContact.UserId;
            if (facebookUserId != -1 && facebookContacts.ContainsKey(facebookUserId))
            {
                _matchedOutlookContacts.Add(new MatchedUserViewModel(outlookContact, facebookContacts[facebookUserId], false) { SynchroniseMatch = true });
                outlookContacts.RemoveAt(index--);
                facebookContacts.Remove(facebookUserId);
            }
            else
            {
                var matchingFacebookContact = facebookContacts.Values.FirstOrDefault(c => c.Name == outlookContact.Name);

                if (matchingFacebookContact != null)
                {
                    outlookContact.UserId = matchingFacebookContact.UserId;
                    _matchedOutlookContacts.Add(new MatchedUserViewModel(outlookContact, matchingFacebookContact, true) { SynchroniseMatch = true });
                    facebookContacts.Remove(matchingFacebookContact.UserId);
                    outlookContacts.RemoveAt(index--);
                }
            }
        }

        private void GetContacts(out IList<IOutlookFacebookUser> outlookContacts, out Dictionary<long, IFacebookUser> facebookContacts)
        {
            //Fetch Outlook and facebook users concurrently but only return from this method after both are done using EventWaitHandles
            IList<IOutlookFacebookUser> olContacts = null;
            Dictionary<long, IFacebookUser> fbContacts = null;

            var outlookResultsAction = ((Action)(() =>
            {
                    olContacts = _outlookRepository.GetContacts();
            }));

            var facebookResultsAction = ((Action)(() =>
            {
                    fbContacts = _facebookRepository.GetFriends().ToDictionary(f => f.UserId);
            }));

            var outlookResults = outlookResultsAction.BeginInvoke(null, null);
            var facebookResults = facebookResultsAction.BeginInvoke(null, null);

            facebookResultsAction.EndInvoke(facebookResults);
            outlookResultsAction.EndInvoke(outlookResults);

            outlookContacts = olContacts;
            facebookContacts = fbContacts;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                if (_contactLoader != null)
                {
                    _contactLoader.Dispose();
                    _contactLoader = null;
                }
            }

            _disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
