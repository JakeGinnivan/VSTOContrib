using System;
using System.Threading;
using System.Windows.Threading;
using FacebookToOutlook.Core;
using FacebookToOutlook.Data;
using FacebookToOutlook.Data.Adapters;
using log4net;
using Outlook.Utility.Services;

namespace FacebookToOutlook
{
    public class FacebookEventSynchronisationService : ISynchronisationService
    {
        private readonly ILog _log;
        private readonly Dispatcher _staDispatcher;
        private int _syncInProgress;
        private bool _syncAgain;
        private readonly GenericSynchronisationService<IFacebookEvent, long> _syncService;
        public event EventHandler<SynchronisationCompleteEventArgs> SynchronisationComplete;

        public FacebookEventSynchronisationService(
            ILog log, 
            IOutlookRepository localRepository,
            IFacebookRepository remoteRepository,
            ISyncSettings syncSettings,
            Dispatcher staDispatcher)
        {
            _log = log;
            _staDispatcher = staDispatcher;

            _syncService = new GenericSynchronisationService<IFacebookEvent, long>(
                 e => e.EventId, 
                 new OutlookRepositorySyncProviderAdapter(localRepository),
                 new FacebookRepositorySyncProviderAdapter(remoteRepository), 
                 syncSettings,
                 SyncDirection.Download);
        }

        public bool SyncInProgress
        {
            get
            {
                return _syncInProgress == 1;
            }
        }

        public void Synchronise()
        {
            //Because Interlocked.CompareExchange has no boolean overload we are using bit's...
            if (Interlocked.CompareExchange(ref _syncInProgress, 1, 1) == 1)
            {
                _syncAgain = true;
                _log.Debug("Sync queued");
                return;
            }

            var done = false;
            while (!done)
            {
                _syncAgain = false;

                var statistics = _syncService.PerformSynchronisation();
                _log.Info(statistics);
                if (SynchronisationComplete != null)
                    _staDispatcher.Invoke(SynchronisationComplete, this, new SynchronisationCompleteEventArgs(statistics));
                done = !_syncAgain;
            }
        }

        public void SynchroniseAsync()
        {
            new Action(Synchronise).BeginInvoke(null, null);
        }
    }
}
