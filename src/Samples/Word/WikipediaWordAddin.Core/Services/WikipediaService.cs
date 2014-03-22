using System.IO;
using System.Net;
using System.Web;
using Newtonsoft.Json;
using WikipediaWordAddin.Services;

namespace WikipediaWordAddin.Core.Services
{
    public class WikipediaService : IWikipediaService
    {
        public SearchResults Search(string search)
        {
            var request = WebRequest.Create("http://en.wikipedia.org/w/api.php?format=json&action=query&list=search&srsearch=" + HttpUtility.UrlEncode(search.Trim()));

            var response = request.GetResponse();

            string json;

            using (var streamReader = new StreamReader(response.GetResponseStream()))
            {
                json = streamReader.ReadToEnd();
            }

            var textReader = new StringReader(json);

            var jsonReader = new JsonTextReader(textReader);

            var serializer = new JsonSerializer();
            return serializer.Deserialize<SearchResults>(jsonReader);
        }
    }
}