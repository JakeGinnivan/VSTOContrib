using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Dynamic;
using System.Linq;
using System.Threading;
using AutoMapper;
using Facebook;
using FacebookToOutlook.Core;
using FacebookToOutlook.Properties;
using FacebookToOutlookCore.Views;

namespace FacebookToOutlook.Data
{
    public class FacebookRepository : IFacebookRepository
    {
        private readonly IApplicationSettings _settings;
        private const string AppId = "128682041968";
        private readonly SynchronizationContext _synchronisationContext;
        private FacebookClient _client;

        public FacebookRepository(IApplicationSettings settings)
        {
            _synchronisationContext = SynchronizationContext.Current;
            _settings = settings;
        }

        public FacebookClient Client
        {
            get
            {
                if (_client == null)
                {
                    _synchronisationContext.Post(s =>
                                                     {
                                                         var login = new FacebookLoginView(AppId, new[] {"user_events"});
                                                         login.ShowDialog();
                                                         if (login.FacebookOAuthResult != null)
                                                         {
                                                             _client =
                                                                 new FacebookClient(
                                                                     login.FacebookOAuthResult.AccessToken);
                                                         }
                                                     }, null);

                    if (_client == null)
                        throw new InvalidOperationException("Facebook client is null =\\");
                }

                return _client;
            }
        }


        public IList<FacebookEvent> GetEvents()
        {
            var events = new List<FacebookEvent>();

            events.AddRange(FetchEventsForStatus(RsvpStatus.Attending, "attending"));
            events.AddRange(FetchEventsForStatus(RsvpStatus.Unsure, "unsure"));
            events.AddRange(FetchEventsForStatus(RsvpStatus.Declined, "declined"));
            events.AddRange(FetchEventsForStatus(RsvpStatus.NotReplied, "not_replied"));

            AddEventsToCache(events);
            
            return events;
        }

        private IEnumerable<FacebookEvent> FetchEventsForStatus(RsvpStatus status, string rsvpStatus)
        {
            if ((_settings.DownloadTypes & status) == status)
            {
                dynamic parameters = new ExpandoObject();
                parameters.rsvp_status = rsvpStatus;
                parameters.start_time = DateTime.Now;
                var o = Client.Get("/me/events", parameters);
                foreach (var e in o)
                {
                    yield return new FacebookEvent(status)
                    {
                        EndTime = e.end_time,
                        EventId = e.id,
                        Name = e.name,
                        StartTime = e.start_time,
                        Location = e.location
                    };
                }
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
                if (_settings.FacebookEventCache == null)
                {
                    _settings.FacebookEventCache = new StringCollection();
                    _settings.Save();                    
                }
                return
                    _settings
                    .FacebookEventCache
                    .Cast<string>()
                    .Select(c =>c.Split('|'))
                    .ToDictionary(c => Convert.ToInt64(c[0]), c => DateTime.Parse(c[1]));
            }
            set
            {
                _settings.FacebookEventCache.Clear();
                foreach (var cacheItem in value)
                {
                    _settings.FacebookEventCache.Add(string.Concat(cacheItem.Key, "|", cacheItem.Value));
                }
                _settings.Save();
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
