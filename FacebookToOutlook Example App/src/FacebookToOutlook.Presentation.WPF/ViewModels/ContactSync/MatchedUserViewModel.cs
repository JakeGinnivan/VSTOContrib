using FacebookToOutlook.Core;
using GalaSoft.MvvmLight;

namespace FacebookToOutlook.Presentation.ViewModels.ContactSync
{
    public class MatchedUserViewModel : ViewModelBase
    {
        private readonly IOutlookFacebookUser _outlookContact;
        private readonly IFacebookUser _matchingFacebookContact;
        private bool _synchroniseMatch;
        private bool _newMatch;

        public MatchedUserViewModel(IOutlookFacebookUser outlookContact, IFacebookUser matchingFacebookContact, bool newMatch)
        {
            _outlookContact = outlookContact;
            _matchingFacebookContact = matchingFacebookContact;
            NewMatch = newMatch;
        }

        public bool NewMatch
        {
            get { return _newMatch; }
            private set
            {
                _newMatch = value;
                RaisePropertyChanged("NewMatch");
            }
        }

        public bool SynchroniseMatch
        {
            get
            {
                return _synchroniseMatch;
            }
            set
            {
                _synchroniseMatch = value;
                RaisePropertyChanged("SynchroniseMatch");
            }
        }

        public IOutlookFacebookUser OutlookContact
        {
            get { return _outlookContact; }
        }

        public IFacebookUser MatchingFacebookContact
        {
            get { return _matchingFacebookContact; }
        }
    }
}
