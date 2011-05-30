using System.ComponentModel;

namespace FacebookToOutlook.Core
{
    public enum BusyStatus
    {
        Free,
        Tentative,
        Busy,
        [Description("Out of Office")]
        OutOfOffice
    }
}