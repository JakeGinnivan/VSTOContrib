using System;

namespace FacebookToOutlook.Core
{
    public class FacebookEvent : IFacebookEvent
    {
        private readonly RsvpStatus _rsvpStatus;

        public FacebookEvent(RsvpStatus rsvpStatus)
        {
            _rsvpStatus = rsvpStatus;
        }

        public long EventId { get; set; }
        public string Name { get; set; }
        public string Location { get; set; }
        public string EventType { get; set; }
        public string EventSubType { get; set; }
        public string Host { get; set; }
        public string Tagline { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public DateTime LastModified { get; set; }

        public RsvpStatus RsvpStatus
        {
            get { return _rsvpStatus; }
        }

    }
}