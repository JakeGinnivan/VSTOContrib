using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Windows.Threading;
using AutoMapper;
using Facebook.Rest;
using Facebook.Schema;
using Facebook.Session;
using FacebookToOutlook.Core;

namespace FacebookToOutlook.Data
{
    public class FacebookRepository : IFacebookRepository
    {
        private readonly Dispatcher _staDispatcher;
        private readonly IConfigurationSettings _settings;
        private readonly ISynchronisedEventInfo _synchronisedEventInfo;
        private const string ApiKey = "bc80e1d1531663f45849be1375ba7a7e";
        private DesktopSession _session;
        private Api _api;

        public FacebookRepository(Dispatcher staDispatcher, IConfigurationSettings settings, ISynchronisedEventInfo synchronisedEventInfo)
        {
            _staDispatcher = staDispatcher;
            _settings = settings;
            _synchronisedEventInfo = synchronisedEventInfo;
        }

        public Api Api
        {
            get
            {
                if (_api == null)
                {
                    var initialiseSession = (Func<Api>)(() =>
                                                {
                                                    _session = new DesktopSession(ApiKey, null, null, true,
                                                                                  new List<Enums.ExtendedPermissions>());
                                                    _session.Login();
                                                    return new Api(_session);
                                                });

                    //Because the Facebook framework deals with UI components if a login is required
                    // we much check access before logging in.
                    if (_staDispatcher.CheckAccess())
                        _api = initialiseSession();
                    else
                        _api = (Api)_staDispatcher.Invoke(initialiseSession);
                }

                return _api;
            }
        }

        public IList<IFacebookUser> GetFriends()
        {
            return Api.Friends.GetUserObjects().Select(u => (IFacebookUser)Mapper.Map(u, new FacebookUser())).ToList();
        }

        public IList<FacebookEvent> GetEvents()
        {
            var events = new List<FacebookEvent>();

            FetchEventsForStatus(events, RsvpStatus.Attending, "attending");
            FetchEventsForStatus(events, RsvpStatus.Unsure, "unsure");
            FetchEventsForStatus(events, RsvpStatus.Declined, "declined");
            FetchEventsForStatus(events, RsvpStatus.NotReplied, "not_replied");

            AddEventsToCache(events);
            
            return events;
        }

        private void FetchEventsForStatus(List<FacebookEvent> events, RsvpStatus status, string rsvpStatus)
        {
            if ((_settings.EventConfigurationSettings.DownloadTypes & status) == status)
            {
                events.AddRange(
                    Api
                        .Events
                        .Get(null, null, DateTime.Now, null, rsvpStatus)
                        .Select(e => Mapper.Map(e, new FacebookEvent(status))));
            }
        }

        /// <summary>
        /// Keeps a list of events and their end times. For calulating deletes
        /// </summary>
        /// <param name="events"></param>
        private void AddEventsToCache(IEnumerable<FacebookEvent> events)
        {
            var cache = EventCache;
            foreach (var facebookEvent in events.Where(facebookEvent => !cache.ContainsKey(facebookEvent.EventId)))
                cache.Add(facebookEvent.EventId, facebookEvent.EndTime);
            EventCache = cache;
        }

        public IList<FacebookEvent> GetModifiedEvents(DateTime since)
        {
            return GetEvents();//.Where(e => e.LastModified >= since).ToList();
        }

        private IDictionary<long, DateTime> EventCache
        {
            get
            {
                if (_synchronisedEventInfo.FacebookEventCache == null)
                {
                    _synchronisedEventInfo.FacebookEventCache = new StringCollection();
                    _synchronisedEventInfo.Save();                    
                }
                return
                    _synchronisedEventInfo
                    .FacebookEventCache
                    .Cast<string>()
                    .Select(c =>c.Split('|'))
                    .ToDictionary(c => Convert.ToInt64(c[0]), c => DateTime.Parse(c[1]));
            }
            set
            {
                _synchronisedEventInfo.FacebookEventCache.Clear();
                foreach (var cacheItem in value)
                {
                    _synchronisedEventInfo.FacebookEventCache.Add(string.Concat(cacheItem.Key, "|", cacheItem.Value));
                }
                _synchronisedEventInfo.Save();
            }
        }

        public IList<long> GetDeletedEventIds()
        {
            var currentEvents = GetEvents();

            return (from long eventCacheKey in EventCache.Keys
                    let cachedEventId = eventCacheKey
                    where EventCache[cachedEventId] >= DateTime.Now
                    where currentEvents.FirstOrDefault(e => e.EventId == cachedEventId) == null
                    select cachedEventId).ToList();
        }
    }
}
