using System.Collections.Generic;

namespace FacebookToOutlook.Services
{
    public interface IOutlookMetaService
    {
        IList<string> GetCategories();
    }
}