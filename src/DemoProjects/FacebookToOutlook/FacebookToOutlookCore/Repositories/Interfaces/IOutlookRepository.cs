using System;
using System.Collections.Generic;
using FacebookToOutlookCore.Model.Interfaces;

namespace FacebookToOutlookCore.Repositories.Interfaces
{
    public interface IOutlookRepository
    {
        IList<IOutlookFacebookUser> GetContacts();
        IList<IOutlookFacebookEvent> GetEvents();
        IList<IOutlookFacebookEvent> GetModifiedEvents(DateTime since);
        IList<long> GetDeletedEventIds();
        bool SaveContacts(IList<IFacebookUser> facebookUsers);
        bool SaveOutlookContacts(IEnumerable<IOutlookFacebookUser> outlookContacts);
        bool SaveEvents(IEnumerable<IFacebookEvent> facebookEvents);
        bool DeleteEvent(long facebookEventId);
        void AssociateFacebookUserWithContact(IOutlookFacebookUser outlookContact, IFacebookUser facebookUserToMatch);
        IOutlookFacebookUser CreateContactFromFacebookUser(IFacebookUser facebookUser);
    }
}
