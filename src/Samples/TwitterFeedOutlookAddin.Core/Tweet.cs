using Newtonsoft.Json;

namespace TwitterFeedOutlookAddin.Core
{
    public class Tweet
    {
        [JsonProperty("user")]
        public User User { get; set; }

        [JsonProperty("id_str")]
        public string Id { get; set; }

        [JsonProperty("text")]
        public string Text { get; set; }
    }
}