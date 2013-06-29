using WikipediaWordAddin.Services;

namespace WikipediaWordAddin.Core.Services
{
    public interface IWikipediaService
    {
        SearchResults Search(string search);
    }
}