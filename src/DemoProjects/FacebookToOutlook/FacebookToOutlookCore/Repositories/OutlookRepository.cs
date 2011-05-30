using System;
using System.Collections.Generic;
using System.Linq;
using AutoMapper;
using FacebookToOutlook.Core;
using FacebookToOutlook.Data.Adapters;
using Microsoft.Office.Interop.Outlook;
using Office.Contrib.Extensions;
using Office.Outlook.Contrib.Extensions;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Outlook.Extensions.Proxies;

namespace FacebookToOutlook.Data
{
    public class OutlookRepository : IOutlookRepository
    {
        private readonly NameSpace _session;
        private readonly IEventConfigurationSettings _settings;

        public OutlookRepository(NameSpace session, IEventConfigurationSettings settings)
        {
            _session = session;
            _settings = settings;
            //Verify user properties are set
            using (var calendar = _session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar).WithComCleanupProxy())
            using (var userProperties = calendar.UserDefinedProperties.WithComCleanupProxy())
            using (var facebookEventIdProperty = userProperties.Find(FacebookEventAdapter.FacebookeventidProperty).WithComCleanupProxy())
            {
                if (facebookEventIdProperty == null)
                    userProperties.Add(FacebookEventAdapter.FacebookeventidProperty, OlUserPropertyType.olText).ReleaseComObject();
            }
        }

        public IList<IOutlookFacebookUser> GetContacts()
        {
            var events = new List<IOutlookFacebookUser>();
            using (var contacts = _session.GetDefaultFolder(OlDefaultFolders.olFolderContacts).WithComCleanup())
            using (var items = contacts.Items.WithComCleanup())
            {
                events.AddRange(
                    items
                        .ComLinq<ContactItem>()
                        .Select(contact => new FacebookUserAdapter(contact))
                        .Select(adapter => Mapper.Map(adapter, new OutlookFacebookUser(adapter.EntryId)))
                        .Cast<IOutlookFacebookUser>());
            }

            return events;
        }

        public IList<IOutlookFacebookEvent> GetEvents()
        {
            var events = new List<IOutlookFacebookEvent>();
            using (var calendar = _session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar).WithComCleanup())
            using (var items = calendar.Items.WithComCleanup())
            {
                events.AddRange(
                    items
                        .ComLinq<AppointmentItem>()
                        .Select(appointment => new FacebookEventAdapter(appointment))
                        .Select(adapter => Mapper.Map(adapter, new OutlookFacebookEvent(adapter.RsvpStatus, adapter.EntryId)))
                        .Where(e => e.EventId != -1)
                        .Cast<IOutlookFacebookEvent>());
            }

            return events;
        }

        public IList<IOutlookFacebookEvent> GetModifiedEvents(DateTime modifiedAfter)
        {
            var events = new List<IOutlookFacebookEvent>();
            using (var calendar = _session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar).WithComCleanupProxy())
            using (var items = calendar.Items.WithComCleanupProxy())
            {
                var lastModStr = modifiedAfter < DateTime.Now.AddMonths(-3)
                                        ?
                                            DateTime.Now.AddMonths(-3).ToString("d/MM/yyy h:mmtt")
                                        :
                                            modifiedAfter.ToString("d/MM/yyy h:mmtt");

                var restrictedItems = items.Restrict("[LastModificationTime] > '" + lastModStr + "'");

                using (var modifiedItems = restrictedItems.WithComCleanupProxy())
                {
                    events.AddRange(
                        modifiedItems
                            .ComLinq<AppointmentItem>()
                            .Select(appointment => new FacebookEventAdapter(appointment))
                            .Select(adapter => Mapper.Map(adapter, new OutlookFacebookEvent(adapter.RsvpStatus, adapter.EntryId)))
                            .Where(e => e.EventId != -1)
                            .Cast<IOutlookFacebookEvent>());
                }
            }

            return events;
        }

        public IList<long> GetDeletedEventIds()
        {
            throw new NotImplementedException();
        }

        public bool SaveOutlookContacts(IEnumerable<IOutlookFacebookUser> outlookContacts)
        {
            using (var contacts = _session.GetDefaultFolder(OlDefaultFolders.olFolderContacts).WithComCleanup())
            using (var items = contacts.Items.WithComCleanup())
            {
                foreach (var outlookContact in outlookContacts)
                {
                    var contactItem = _session.GetItemFromID(outlookContact.EntryId, contacts.StoreID) as _ContactItem;
                    using (var outlookContactItem = contactItem.WithComCleanup())
                    {
                        CreateOrUpdateContact(outlookContact, outlookContactItem, items);
                    }
                }
            }

            return true;
        }

        public bool SaveContacts(IList<IFacebookUser> facebookUsers)
        {
            using (var contacts = _session.GetDefaultFolder(OlDefaultFolders.olFolderContacts).WithComCleanup())
            using (var items = contacts.Items.WithComCleanup())
            {
                foreach (var facebookContact in facebookUsers)
                {
                    var filter = string.Format("[{0}] = '{1}'", FacebookUserAdapter.FacebookUserIdProperty, facebookContact.UserId);
                    var contactItem = items.Find(filter) as _ContactItem;
                    using (var outlookContact = contactItem.WithComCleanup())
                    {
                        CreateOrUpdateContact(facebookContact, outlookContact, items);
                    }
                }
            }

            return true;
        }

        private static void CreateOrUpdateContact(IFacebookUser facebookContact, _ContactItem outlookContact, _Items items)
        {
            if (outlookContact != null)
            {
                UpdateAdapter(facebookContact, outlookContact);
            }
            else using (var newItem = ((_ContactItem)items.Add(OlItemType.olContactItem)).WithComCleanup())
                {
                    UpdateAdapter(facebookContact, newItem);
                }
        }

        private static void UpdateAdapter(IFacebookUser facebookUser, _ContactItem outlookContact)
        {
            var itemAdapter = new FacebookUserAdapter(outlookContact);
            Mapper.Map(facebookUser, itemAdapter);

            if (!outlookContact.Saved)
                outlookContact.Save();
        }

        public bool SaveEvents(IEnumerable<IFacebookEvent> facebookEvents)
        {
            using (var calendar = _session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar).WithComCleanup())
            using (var items = calendar.Items.WithComCleanup())
            {
                foreach (var facebookEvent in facebookEvents)
                {
                    SaveEvent(items, facebookEvent);
                }
            }

            return true;
        }
        
        private void SaveEvent(_Items items, IFacebookEvent facebookEvent)
        {
            var filter = string.Format("[{0}] = '{1}'", FacebookEventAdapter.FacebookeventidProperty, facebookEvent.EventId);
            using (var outlookAppointment = (items.Find(filter) as _AppointmentItem).WithComCleanup())
            {
                if (outlookAppointment != null)
                {
                    UpdateAdapter(facebookEvent, outlookAppointment);
                }
                else using (var newItem = ((_AppointmentItem)items.Add(OlItemType.olAppointmentItem)).WithComCleanup())
                {
                    UpdateAdapter(facebookEvent, newItem);
                }
            }
        }

        private void UpdateAdapter(IFacebookEvent facebookEvent, _AppointmentItem outlookAppointment)
        {
            var itemAdapter = new FacebookEventAdapter(outlookAppointment);
            Mapper.Map(facebookEvent, itemAdapter);
            ApplySettings(outlookAppointment);
            outlookAppointment.Save();
        }

        private void ApplySettings(_AppointmentItem newItem)
        {
            if (_settings.MarkAsPrivate)
                newItem.Sensitivity = OlSensitivity.olPrivate;

            if (_settings.EventReminder)
            {
                newItem.ReminderSet = true;
                newItem.ReminderMinutesBeforeStart = _settings.RemindMinutesBefore;
            }

            newItem.Categories = _settings.Category;

            switch (_settings.ShowTimeAs)
            {
                case BusyStatus.Free:
                    newItem.BusyStatus = OlBusyStatus.olFree;
                    break;
                case BusyStatus.Tentative:
                    newItem.BusyStatus = OlBusyStatus.olTentative;
                    break;
                case BusyStatus.Busy:
                    newItem.BusyStatus = OlBusyStatus.olBusy;
                    break;
                case BusyStatus.OutOfOffice:
                    newItem.BusyStatus = OlBusyStatus.olOutOfOffice;
                    break;
            }
        }

        public bool DeleteEvent(long facebookEventId)
        {
            var filter = string.Format("[{0}] = '{1}'", FacebookEventAdapter.FacebookeventidProperty, facebookEventId);
            using (var calendar = _session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar).WithComCleanup())
            using (var items = calendar.Items.WithComCleanup())
            using (var outlookAppointment = (items.Find(filter) as _AppointmentItem).WithComCleanup())
            {
                if (outlookAppointment != null)
                    outlookAppointment.Delete();
            }

            return true;
        }

        public void AssociateFacebookUserWithContact(IOutlookFacebookUser outlookContact, IFacebookUser facebookUserToMatch)
        {
            using (var contacts = _session.GetDefaultFolder(OlDefaultFolders.olFolderContacts).WithComCleanup())
            using (var contact = (_session.GetItemFromID(outlookContact.EntryId, contacts.StoreID) as _ContactItem).WithComCleanup())
            {
                new FacebookUserAdapter(contact) { UserId = facebookUserToMatch.UserId };
                contact.Save();
            }
        }

        public IOutlookFacebookUser CreateContactFromFacebookUser(IFacebookUser facebookUser)
        {
            using (var contactFolder = _session.GetDefaultFolder(OlDefaultFolders.olFolderContacts).WithComCleanup())
            using (var contactItems = contactFolder.Items.WithComCleanup())
            using (var contact = (contactItems.Add(OlItemType.olContactItem) as ContactItem).WithComCleanup())
            {
                var userAdapter = new FacebookUserAdapter(contact);
                Mapper.Map(facebookUser, userAdapter);

                contact.Display(true);

                return contact.Saved ? Mapper.Map(userAdapter, new OutlookFacebookUser(contact.EntryID)) : null;
            }
        }
    }
}
