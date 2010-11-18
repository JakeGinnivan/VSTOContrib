using System.Collections.Generic;
using System.IO;
using System.Net;
using Newtonsoft.Json;
using TwitterFeedCore.TwitterFeed;

namespace TwitterFeedCore.Services
{
    public class TwitterService : ITwitterService
    {
        public List<Tweet> GetTwitterStreamForUsername(string username)
        {
            var request = WebRequest.Create("http://api.twitter.com/1/statuses/user_timeline.json?count=200&screen_name=" + username);

            var response = request.GetResponse();

            string json = string.Empty;

            using (var streamReader = new StreamReader(response.GetResponseStream()))
            {
                json = streamReader.ReadToEnd();
            }

            var textReader = new StringReader(json);

            var jsonReader = new JsonTextReader(textReader);

            var serializer = new JsonSerializer();
            return serializer.Deserialize<List<Tweet>>(jsonReader);
        }
    }
}
