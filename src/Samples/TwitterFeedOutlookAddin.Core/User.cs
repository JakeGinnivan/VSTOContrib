using Newtonsoft.Json;

namespace TwitterFeedOutlookAddin.Core
{
    public class User
    {
        [JsonProperty("screen_name")]
        public string ScreenName { get; set; }
    }
}