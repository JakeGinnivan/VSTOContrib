using System;
using System.Collections.Generic;
using System.Windows.Threading;
using FacebookToOutlook.Core;
using Microsoft.Office.Interop.Outlook;

namespace FacebookToOutlook.Data
{
    public class OutlookDispatchingRepository : IOutlookRepository
    {
        private readonly Dispatcher _outlookStaDispatcher;
        private readonly OutlookRepository _outlookEventRepository;

        public OutlookDispatchingRepository(Dispatcher outlookStaDispatcher, NameSpace session, IEventConfigurationSettings settings)
        {
            _outlookEventRepository = new OutlookRepository(session, settings);
            _outlookStaDispatcher = outlookStaDispatcher;
        }

        public IList<IOutlookFacebookUser> GetContacts()
        {
            var getContacts = ((Func<IList<IOutlookFacebookUser>>)(() => _outlookEventRepository.GetContacts()));

            return (IList<IOutlookFacebookUser>)_outlookStaDispatcher.Invoke(getContacts);
        }

        public IList<IOutlookFacebookEvent> GetEvents()
        {
            var getEvents = ((Func<IList<IOutlookFacebookEvent>>)(() => _outlookEventRepository.GetEvents()));

            return (IList<IOutlookFacebookEvent>)_outlookStaDispatcher.Invoke(getEvents);
        }

        public IList<IOutlookFacebookEvent> GetModifiedEvents(DateTime since)
        {
            var getEvents = ((Func<IList<IOutlookFacebookEvent>>)(() => _outlookEventRepository.GetModifiedEvents(since)));

            return (IList<IOutlookFacebookEvent>)_outlookStaDispatcher.Invoke(getEvents);
        }

        public IList<long> GetDeletedEventIds()
        {
            var getEvents = ((Func<IList<long>>)(() => _outlookEventRepository.GetDeletedEventIds()));

            return (IList<long>)_outlookStaDispatcher.Invoke(getEvents);
        }

        public bool SaveContacts(IList<IFacebookUser> facebookUsers)
        {
            var getEvents = ((Func<IList<IFacebookUser>, bool>)(contacts => _outlookEventRepository.SaveContacts(contacts)));

            return ((bool?)_outlookStaDispatcher.Invoke(getEvents, facebookUsers)) ?? false;
        }

        public bool SaveEvents(IEnumerable<IFacebookEvent> facebookEvents)
        {
            var getEvents = ((Func<IEnumerable<IFacebookEvent>, bool>)(events => _outlookEventRepository.SaveEvents(events)));

            return ((bool?) _outlookStaDispatcher.Invoke(getEvents, facebookEvents))??false;
        }

        public bool DeleteEvent(long facebookEventId)
        {
            var getEvents = ((Func<long,bool>)(eventId => _outlookEventRepository.DeleteEvent(eventId)));

            return ((bool?) _outlookStaDispatcher.Invoke(getEvents, facebookEventId)) ?? false;
        }

        public void AssociateFacebookUserWithContact(IOutlookFacebookUser outlookContact, IFacebookUser facebookUserToMatch)
        {
            var getEvents = ((Action<IOutlookFacebookUser, IFacebookUser>)((oContact, fbUserToMatch) => _outlookEventRepository.AssociateFacebookUserWithContact(oContact, fbUserToMatch)));

            _outlookStaDispatcher.Invoke(getEvents, outlookContact, facebookUserToMatch);
        }

        public IOutlookFacebookUser CreateContactFromFacebookUser(IFacebookUser facebookUser)
        {
            var createContactFromFacebookUserAction = ((Func<IFacebookUser, IOutlookFacebookUser>)(fbUser => _outlookEventRepository.CreateContactFromFacebookUser(fbUser)));

            return ((IOutlookFacebookUser)_outlookStaDispatcher.Invoke(createContactFromFacebookUserAction, facebookUser));
        }

        public bool SaveOutlookContacts(IEnumerable<IOutlookFacebookUser> outlookContacts)
        {
            var getEvents = ((Func<IList<IOutlookFacebookUser>, bool>)(contacts => _outlookEventRepository.SaveOutlookContacts(contacts)));

            return ((bool?)_outlookStaDispatcher.Invoke(getEvents, outlookContacts)) ?? false;
        }
    }
}
