using System.Collections.Generic;

namespace WikipediaWordAddin.Services
{
    public class Query
    {
        public SearchInfo searchinfo { get; set; }
        public List<SearchResult> search { get; set; }
    }
}