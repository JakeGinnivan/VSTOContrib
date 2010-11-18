using Newtonsoft.Json;

namespace TwitterFeedCore.TwitterFeed
{
    public class User
    {
        [JsonProperty("screen_name")]
        public string ScreenName { get; set; }
    }
}
