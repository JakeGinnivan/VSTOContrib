using System.Collections.Generic;
using TwitterFeedCore.TwitterFeed;

namespace TwitterFeedCore.Services
{
    public interface ITwitterService
    {
        List<Tweet> GetTwitterStreamForUsername(string username);
    }
}