using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Outlook;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Outlook.Extensions.Proxies;

namespace FacebookToOutlook.Services
{
    public class OutlookMetaService : IOutlookMetaService
    {
        private readonly NameSpace _session;

        public OutlookMetaService(NameSpace session)
        {
            _session = session;
        }

        public IList<string> GetCategories()
        {
            using (var categories = _session.Categories.WithComCleanupProxy())
            {
                return categories.ComLinq<Category>().Select(c=>c.Name).ToList();
            }
        }
    }
}
