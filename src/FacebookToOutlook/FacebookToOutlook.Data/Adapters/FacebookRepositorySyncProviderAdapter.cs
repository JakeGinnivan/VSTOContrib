using System;
using System.Collections.Generic;
using FacebookToOutlook.Core;
using Outlook.Utility.Services;

namespace FacebookToOutlook.Data.Adapters
{
    public class FacebookRepositorySyncProviderAdapter : ISynchronisationProvider<IFacebookEvent, long>
    {
        private readonly IFacebookRepository _facebookRepository;

        public FacebookRepositorySyncProviderAdapter(IFacebookRepository facebookRepository)
        {
            _facebookRepository = facebookRepository;
        }

        public IEnumerable<IFacebookEvent> GetModifiedEntries(DateTime? lastSync)
        {
            return
                lastSync == null
                    ?
                        _facebookRepository.GetEvents()
                    :
                        _facebookRepository.GetModifiedEvents(lastSync.Value);
        }

        public IEnumerable<long> GetDeletedEntries(DateTime? lastSync)
        {
            return _facebookRepository.GetDeletedEventIds();
        }

        public void SaveEntries(IEnumerable<IFacebookEvent> entries)
        {
            throw new InvalidOperationException("Read only sync provider");
        }

        public void DeleteEntries(IEnumerable<long> keys)
        {
            throw new InvalidOperationException("Read only sync provider");
        }
    }
}
