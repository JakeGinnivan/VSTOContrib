using System;

namespace FacebookToOutlook.Core
{
    [Flags]
    public enum RsvpStatus
    {
        None = 0,
        Attending = 1,
        Unsure = 2,
        Declined = 4,
        NotReplied = 8
    }
}
