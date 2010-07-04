using System.Collections.Specialized;

namespace FacebookToOutlook.Data
{
    public interface ISynchronisedEventInfo
    {
        StringCollection FacebookEventCache { get; set; }
        void Save();
    }
}
