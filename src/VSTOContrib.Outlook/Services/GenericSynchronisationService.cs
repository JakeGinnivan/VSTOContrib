using System;
using System.Collections.Generic;
using System.Linq;
using VSTOContrib.Outlook.Properties;
using VSTOContrib.Outlook.Services.Conficts;

namespace VSTOContrib.Outlook.Services
{
    /// <summary>
    /// Generic Synchronisation Service which takes care of the basic synchronisation logic
    /// and lets you focus on providing the items and conflict resolution
    /// </summary>
    /// <typeparam name="TType">The type of the type.</typeparam>
    /// <typeparam name="TKey">The type of the key.</typeparam>
    public class GenericSynchronisationService<TType, TKey> where TType : class
    {
        private readonly Func<TType, TKey> _keySelector;
        private readonly ISynchronisationProvider<TType, TKey> _localProvider;
        private readonly ISynchronisationProvider<TType, TKey> _remoteProvider;
        private readonly ISyncSettings _settings;
        private readonly SyncDirection _syncDirection;
        private IConflictResolver<TType> _conflcitResolver = new RemoteWinsResolver<TType>();

        /// <summary>
        /// Initializes a new instance of the <see cref="GenericSynchronisationService&lt;TType, TKey&gt;"/> class.
        /// </summary>
        /// <param name="keySelector">The key selector.</param>
        /// <param name="localProvider">The local provider.</param>
        /// <param name="remoteProvider">The remote provider.</param>
        /// <param name="settings">The settings.</param>
        /// <param name="syncDirection">The sync direction.</param>
        public GenericSynchronisationService(
            Func<TType, TKey> keySelector, 
            ISynchronisationProvider<TType, TKey> localProvider, 
            ISynchronisationProvider<TType, TKey> remoteProvider,
            ISyncSettings settings,
            SyncDirection syncDirection)
        {
            _keySelector = keySelector;
            _localProvider = localProvider;
            _remoteProvider = remoteProvider;
            _settings = settings;
            _syncDirection = syncDirection;
        }

        /// <summary>
        /// Gets or sets the conflcit resolver.
        /// </summary>
        /// <value>The conflcit resolver.</value>
        public IConflictResolver<TType> ConflcitResolver
        {
            get { return _conflcitResolver; }
            set
            {
                if (value == null) throw new ArgumentNullException("value", Resources.NoConflictResolverMessage);
                _conflcitResolver = value;
            }
        }

        /// <summary>
        /// Performs the synchronisation.
        /// </summary>
        /// <returns></returns>
        public SynchronisationResults PerformSynchronisation()
        {
            switch(_syncDirection)
            {
                case SyncDirection.Download:
                    return DownloadChanges();
                case SyncDirection.Upload:
                    return UploadChanges();
                case SyncDirection.TwoWay:
                    return TwoWaySync();
                default:
                    throw new ArgumentException("Unrecognised sync direction");
            }
        }

        private SynchronisationResults TwoWaySync()
        {
            List<TType> saveLocal;
            List<TType> saveRemote;
            var remoteDeleted = _remoteProvider.GetDeletedEntries(_settings.LastSync).ToList();
            var localDeleted = _localProvider.GetDeletedEntries(_settings.LastSync).ToList();

            var results = PopulateItemsToSave(
                _remoteProvider.GetModifiedEntries(_settings.LastSync),
                _localProvider.GetModifiedEntries(_settings.LastSync), 
                out saveLocal, out saveRemote);

            try
            {
                _remoteProvider.SaveEntries(saveRemote);
                _localProvider.SaveEntries(saveLocal);
                _localProvider.DeleteEntries(remoteDeleted);
                _remoteProvider.DeleteEntries(localDeleted);
            }
            catch (Exception ex)
            {
                results.Success = false;
                results.Message = ex.Message;
                return results;
            }

            results.NumberDeletedRemote = localDeleted.Count;
            results.NumberDeletedLocal = remoteDeleted.Count;
            results.NumberUpdatedLocal = saveLocal.Count;
            results.NumberUpdatedRemote = saveRemote.Count;

            _settings.LastSync = DateTime.Now;
            _settings.Save();
            results.Success = true;
            return results;
        }

        private SynchronisationResults UploadChanges()
        {
            List<TType> saveLocal;
            List<TType> saveRemote;
            var localDeleted = _localProvider.GetDeletedEntries(_settings.LastSync).ToList();

            var results = PopulateItemsToSave(
                new TType[0], 
                _localProvider.GetModifiedEntries(_settings.LastSync),
                out saveLocal, out saveRemote);

            try
            {
                _remoteProvider.SaveEntries(saveRemote);
                _remoteProvider.DeleteEntries(localDeleted);
            }
            catch (Exception ex)
            {
                results.Success = false;
                results.Message = ex.Message;
                return results;
            }

            results.NumberDeletedRemote = localDeleted.Count;
            results.NumberUpdatedRemote = saveRemote.Count;

            _settings.LastSync = DateTime.Now;
            _settings.Save();
            results.Success = true;
            return results;
        }

        private SynchronisationResults DownloadChanges()
        {
            List<TType> saveLocal;
            List<TType> saveRemote;
            var localDeleted = _remoteProvider.GetDeletedEntries(_settings.LastSync).ToList();

            var results = PopulateItemsToSave(
                _remoteProvider.GetModifiedEntries(_settings.LastSync),
                new TType[0],
                out saveLocal, out saveRemote);

            try
            {
                _localProvider.SaveEntries(saveLocal);
                _localProvider.DeleteEntries(localDeleted);
            }
            catch (Exception ex)
            {
                results.Success = false;
                results.Message = ex.Message;
                return results;
            }

            results.NumberDeletedLocal = localDeleted.Count;
            results.NumberUpdatedLocal = saveLocal.Count;
            results.NumberDeletedRemote = 0;
            results.NumberUpdatedRemote = saveRemote.Count;

            _settings.LastSync = DateTime.Now;
            _settings.Save();
            results.Success = true;
            return results;
        }

        private SynchronisationResults PopulateItemsToSave(
            IEnumerable<TType> remoteModifications, IEnumerable<TType> localModifications, 
            out List<TType> saveLocal, out List<TType> saveRemote)
        {
            var remoteModified = remoteModifications.ToDictionary(_keySelector);
            var localModified = localModifications.ToDictionary(_keySelector);
            var results = new SynchronisationResults();

            saveLocal = new List<TType>();
            saveRemote = new List<TType>();
            foreach (var remoteEntry in remoteModified)
            {
                var entry = remoteEntry;
                var localMatchingModifications = localModified.Where(t => Equals(t.Key, entry.Key)).Select(e =>e.Value).ToList();

                switch (localMatchingModifications.Count)
                {
                    case 0:
                        saveLocal.Add(remoteEntry.Value);
                        break;
                    case 1:
                        HandleConflict(results, localMatchingModifications[0], remoteEntry.Value,
                            saveLocal, saveRemote);
                        break;
                }
            }

            saveRemote.AddRange(localModified.Where(modifiedEntry => !remoteModified.Keys.Contains(modifiedEntry.Key)).Select(l => l.Value).ToList());
            return results;
        }

        private void HandleConflict(SynchronisationResults results, 
            TType localEntry, TType remoteEntry,
            ICollection<TType> saveLocal, ICollection<TType> saveRemote)
        {
            var resolution = ConflcitResolver.Resolve(localEntry, remoteEntry);
            if (resolution.SaveLocal != null)
            {
                if (_syncDirection == SyncDirection.Upload)
                    throw new InvalidOperationException("Sync direction is upload, but conflict resolved as saving locally.");
                saveLocal.Add(resolution.SaveLocal);
            }
            if (resolution.SaveRemote != null)
            {
                if (_syncDirection == SyncDirection.Download)
                    throw new InvalidOperationException("Sync direction is download, but conflict resolved as saving remotely.");
                saveRemote.Add(resolution.SaveRemote);
            }

            results.NumberConflicts++;
        }
    }
}
