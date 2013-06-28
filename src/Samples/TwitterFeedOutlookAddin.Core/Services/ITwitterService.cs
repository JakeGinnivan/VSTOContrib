using System.Collections.Generic;

namespace TwitterFeedOutlookAddin.Core.Services
{
    public interface ITwitterService
    {
        List<Tweet> GetTwitterStreamForUsername(string username);
    }
}