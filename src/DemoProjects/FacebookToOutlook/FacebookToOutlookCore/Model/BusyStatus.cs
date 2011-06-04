using System.ComponentModel;

namespace FacebookToOutlookCore.Model
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