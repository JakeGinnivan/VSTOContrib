using System.Collections.Generic;

namespace FacebookToOutlookCore.Services.Interfaces
{
    public interface IOutlookMetaService
    {
        IList<string> GetCategories();
    }
}