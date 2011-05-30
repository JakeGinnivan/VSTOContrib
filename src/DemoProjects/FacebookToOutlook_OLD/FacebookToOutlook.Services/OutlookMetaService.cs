using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Outlook;
using Office.Contrib.Extensions;
using Office.Outlook.Contrib.Extensions;

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
            using (var categories = _session.Categories.WithComCleanup())
            {
                return categories.ComLinq<Category>().Select(c=>c.Name).ToList();
            }
        }
    }
}
