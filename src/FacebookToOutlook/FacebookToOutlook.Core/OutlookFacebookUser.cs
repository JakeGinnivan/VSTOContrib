namespace FacebookToOutlook.Core
{
    public class OutlookFacebookUser : FacebookUser, IOutlookFacebookUser
    {
        private readonly string _entryId;

        public OutlookFacebookUser(string entryId)
        {
            _entryId = entryId;
        }

        public string EntryId
        {
            get { return _entryId; }
        }
    }
}