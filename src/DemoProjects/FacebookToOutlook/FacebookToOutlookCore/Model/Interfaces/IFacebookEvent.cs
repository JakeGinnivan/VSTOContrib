using System;

namespace FacebookToOutlookCore.Model.Interfaces
{
    public interface IFacebookEvent
    {
        long EventId { get; set; }
        string Name { get; set; }
        string Location { get; set; }
        string EventType { get; set; }
        string EventSubType { get; set; }
        string Host { get; set; }
        string Tagline { get; set; }
        DateTime StartTime { get; set; }
        DateTime EndTime { get; set; }
        DateTime LastModified { get; }
        RsvpStatus RsvpStatus { get; }
    }
}