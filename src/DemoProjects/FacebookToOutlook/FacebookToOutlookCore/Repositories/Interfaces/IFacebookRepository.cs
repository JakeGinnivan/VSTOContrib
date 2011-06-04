using System;
using System.Collections.Generic;
using FacebookToOutlookCore.Model;

namespace FacebookToOutlookCore.Repositories.Interfaces
{
    public interface IFacebookRepository
    {
        IList<FacebookEvent> GetEvents();
        IList<FacebookEvent> GetModifiedEvents(DateTime since);
        IList<long> GetDeletedEventIds();
    }
}
