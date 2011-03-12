using System;
using System.Text;
using FacebookToOutlook.Core;
using Microsoft.Office.Interop.Outlook;
using Office.Outlook.Contrib.Extensions;

namespace FacebookToOutlook.Data.Adapters
{
    public class FacebookEventAdapter : IOutlookFacebookEvent
    {
        public const string FacebookeventidProperty = "FacebookEventId";
        public const string RsvpStatusProperty = "FacebookRsvp";
        public const string IsFacebookEventProperty = "IsFacebookEvent";
        private readonly _AppointmentItem _appointmentItem;
        private string _eventType;
        private string _eventSubType;
        private string _tagline;
        private string _host;

        public FacebookEventAdapter(_AppointmentItem appointmentItem)
        {
            _appointmentItem = appointmentItem;
        }

        public FacebookEventAdapter(_AppointmentItem appointmentItem, RsvpStatus rsvpStatus)
        {
            _appointmentItem = appointmentItem;
            RsvpStatus = rsvpStatus;
        }

        public string EntryId
        {
            get { return _appointmentItem.EntryID; }
        }

        public long EventId
        {
            get
            {
                return _appointmentItem.GetPropertyValue(FacebookeventidProperty, OlUserPropertyType.olText, false, Convert.ToInt64, -1);
            }
            set
            {
                _appointmentItem.SetPropertyValue(FacebookeventidProperty, OlUserPropertyType.olText, value.ToString(), true);
                IsFacebookEvent = (value != -1);
            }
        }

        public bool IsFacebookEvent
        {
            get
            {
                return _appointmentItem.GetPropertyValue(IsFacebookEventProperty, OlUserPropertyType.olYesNo, false, Convert.ToBoolean, false);
            }
            private set
            {
                _appointmentItem.SetPropertyValue(IsFacebookEventProperty, OlUserPropertyType.olYesNo, value, true);                
            }
        }

        public string Name
        {
            get
            {
                return _appointmentItem.Subject;
            }
            set
            {
                _appointmentItem.Subject = value;
            }
        }

        public string Location
        {
            get
            {
                return _appointmentItem.Location;
            }
            set
            {
                _appointmentItem.Location = value;
            }
        }


        public string EventType
        {
            get { return _eventType; }
            set
            {
                _eventType = value;
                UpdateBody();
            }
        }

        public string EventSubType
        {
            get { return _eventSubType; }
            set
            {
                _eventSubType = value;
                UpdateBody();
            }
        }

        public string Tagline
        {
            get { return _tagline; }
            set
            {
                _tagline = value;
                UpdateBody();
            }
        }

        public string Host
        {
            get { return _host; }
            set
            {
                _host = value;
                UpdateBody();
            }
        }

        private void UpdateBody()
        {
            var bodyBuilder = new StringBuilder();

            if (!string.IsNullOrEmpty(Tagline))
            {
                bodyBuilder.AppendLine(Tagline);
                bodyBuilder.AppendLine();
            }
            if (!string.IsNullOrEmpty(Host))
                bodyBuilder.AppendLine(string.Format("Host: {0}", Host));
            if (!string.IsNullOrEmpty(EventType))
                bodyBuilder.AppendLine(string.Format("Event Type: {0}", EventType));
            if (!string.IsNullOrEmpty(EventSubType))
                bodyBuilder.AppendLine(string.Format("Event Sub Type: {0}", EventSubType));

            _appointmentItem.Body = bodyBuilder.ToString();
        }

        public DateTime StartTime
        {
            get
            {
                return _appointmentItem.Start;
            }
            set
            {
                _appointmentItem.Start = value;
            }
        }

        public DateTime EndTime
        {
            get
            {
                return _appointmentItem.End;
            }
            set
            {
                _appointmentItem.End = value;
            }
        }

        public DateTime LastModified
        {
            get
            {
                return _appointmentItem.LastModificationTime;
            }
        }

        public RsvpStatus RsvpStatus
        {
            get
            {
                var propertyValue = _appointmentItem.GetPropertyValue(RsvpStatusProperty, OlUserPropertyType.olText, false, Convert.ToString, RsvpStatus.None.ToString());
                return (RsvpStatus) Enum.Parse(typeof(RsvpStatus), propertyValue);
            }
            private set
            {
                _appointmentItem.SetPropertyValue(RsvpStatusProperty, OlUserPropertyType.olText, value.ToString(), true);
            }
        }
    }
}
