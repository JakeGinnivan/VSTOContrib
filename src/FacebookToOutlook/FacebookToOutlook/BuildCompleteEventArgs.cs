using System;
using System.Collections.ObjectModel;
using FacebookToOutlook.Presentation.ViewModels;
using FacebookToOutlook.Presentation.ViewModels.ContactSync;

namespace FacebookToOutlook.Presentation
{
    public class BuildCompleteEventArgs : EventArgs
    {
        private readonly bool _success;
        private readonly string _errorMessage;
        private readonly ObservableCollection<MatchedUserViewModel> _matchedUsers;
        private readonly UnmatchedContactsViewModel _unmatchedList;

        public BuildCompleteEventArgs(bool success, string errorMessage, ObservableCollection<MatchedUserViewModel> matchedUsers, UnmatchedContactsViewModel unmatchedList)
        {
            _success = success;
            _errorMessage = errorMessage;
            _matchedUsers = matchedUsers;
            _unmatchedList = unmatchedList;
        }

        public bool Success
        {
            get { return _success; }
        }

        public string ErrorMessage
        {
            get { return _errorMessage; }
        }

        public ObservableCollection<MatchedUserViewModel> MatchedUsers
        {
            get { return _matchedUsers; }
        }

        public UnmatchedContactsViewModel UnmatchedList
        {
            get { return _unmatchedList; }
        }
    }
}