using FacebookToOutlookCore.Model.Interfaces;

namespace FacebookToOutlookCore.Model
{
    public class OutlookFacebookEvent : FacebookEvent, IOutlookFacebookEvent
    {
        private readonly string _eventId;

        public OutlookFacebookEvent(RsvpStatus rsvpStatus, string eventId) : base(rsvpStatus)
        {
            _eventId = eventId;
        }

        public string EntryId
        {
            get { return _eventId; }
        }
    }
}