using System;

namespace FacebookToOutlook.Presentation.ViewModels.ContactSync
{
    public class NewMatchEventArgs : EventArgs
    {
        private readonly MatchedUserViewModel _newMatch;

        public NewMatchEventArgs(MatchedUserViewModel newMatch)
        {
            _newMatch = newMatch;
        }

        public MatchedUserViewModel NewMatch
        {
            get { return _newMatch; }
        }
    }
}