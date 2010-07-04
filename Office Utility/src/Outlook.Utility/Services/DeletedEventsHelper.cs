using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;

namespace Outlook.Utility.Services
{
    /// <summary>
    /// Helper class which helps you track deletions of events (entities that have an end date).
    /// Often remote systems will not return events from the past, we do not want these
    /// events to be deleted
    /// </summary>
    /// <typeparam name="TEntity">Entity to track deletions for</typeparam>
    /// <typeparam name="TKey">Id for the Entity</typeparam>
    public class DeletedEventsHelper<TEntity, TKey> 
        where TKey : struct 
        where TEntity : class
    {
        private readonly Func<TEntity, TKey> _keySelector;
        private readonly Func<TEntity, DateTime> _endDateSelector;
        private readonly SettingsStore _store;
        private Dictionary<TKey, DateTime> _keyCache;

        /// <summary>
        /// Initializes a new instance of the <see cref="DeletedEventsHelper&lt;TEntity, TKey&gt;"/> class.
        /// </summary>
        /// <param name="keySelector">The key selector.</param>
        /// <param name="endDateSelector">The end date selector.</param>
        public DeletedEventsHelper(Func<TEntity, TKey> keySelector, Func<TEntity, DateTime> endDateSelector)
        {
            _keySelector = keySelector;
            _endDateSelector = endDateSelector;
            _store = new SettingsStore();
        }

        private IDictionary<TKey, DateTime> KeyCache
        {
            get
            {
                if (_keyCache == null)
                {
                    if (_store.EntityKeyCache == null)
                    {
                        _store.EntityKeyCache = new Hashtable();
                        _store.Save();
                    }

                    _keyCache = new Dictionary<TKey, DateTime>();
                    foreach (var key in _store.EntityKeyCache.Keys)
                    {
                        _keyCache.Add((TKey)key, (DateTime)_store.EntityKeyCache[key]);
                    }
                }
                
                return _keyCache;
            }
            set
            {
                if (_store.EntityKeyCache == null) _store.EntityKeyCache = new Hashtable();
                if (_keyCache == null) _keyCache = new Dictionary<TKey, DateTime>();
                _store.EntityKeyCache.Clear();
                _keyCache.Clear();
                foreach (var cacheItem in value)
                {
                    _store.EntityKeyCache.Add(cacheItem.Key, cacheItem.Value);
                    _keyCache.Add(cacheItem.Key, cacheItem.Value);
                }
                _store.Save();
            }
        }

        /// <summary>
        /// Adds the events to known items cache.
        /// </summary>
        /// <param name="entities">The entities.</param>
        public void AddEventsToCache(IEnumerable<TEntity> entities)
        {
            KeyCache = entities.ToDictionary(e => _keySelector(e), e => _endDateSelector(e));
        }

        /// <summary>
        /// Gets the deleted items.
        /// </summary>
        /// <param name="entities">The entities.</param>
        /// <returns></returns>
        public IList<TKey> GetDeletedItems(IEnumerable<TEntity> entities)
        {
            // Only return Id's that have an end date after current time
            return (from eventCacheKey in KeyCache.Keys
                    let cachedEventId = eventCacheKey
                    where KeyCache[cachedEventId] >= DateTime.Today
                    where entities.FirstOrDefault(e => _keySelector(e).Equals(cachedEventId)) == null
                    select cachedEventId).ToList();
        }

        private class SettingsStore : ApplicationSettingsBase
        {
            public SettingsStore()
                : base(typeof(TEntity).FullName + typeof(TKey).FullName)
            { }

            [UserScopedSetting]
            public Hashtable EntityKeyCache { get; set; }
        }
    }
}
