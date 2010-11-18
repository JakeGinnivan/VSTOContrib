using System.Collections.Generic;
using FacebookToOutlook.Core;
using FacebookToOutlook.Services;

namespace FacebookToOutlook.Presentation.ViewModels
{
    public class EventConfigurationViewModel : ErrorInfoViewModelBase, IEventConfigurationSettings
    {
        private bool _markAsPrivate;
        private bool _eventReminder;
        private int _remindMinutesBefore;
        private BusyStatus _showTimeAs;
        private string _category;
        private RsvpStatus _downloadTypes;
        private readonly IList<string> _outlookCategories;

        public EventConfigurationViewModel(IOutlookMetaService outlookMetaService)
        {
            _outlookCategories = outlookMetaService.GetCategories();
        }

        public IList<string> OutlookCategories
        {
            get { return _outlookCategories; }
        }

        public bool MarkAsPrivate
        {
            get { return _markAsPrivate; }
            set
            {
                _markAsPrivate = value;
                RaisePropertyChanged("MarkAsPrivate");
            }
        }

        public bool EventReminder
        {
            get { return _eventReminder; }
            set
            {
                _eventReminder = value;
                RaisePropertyChanged("EventReminder");
            }
        }

        public int RemindMinutesBefore
        {
            get { return _remindMinutesBefore; }
            set
            {
                const string remindMinutesBeforeProperty = "RemindMinutesBefore";
                if (value < 0)
                {
                    SetError(remindMinutesBeforeProperty, "Reminder must be >= 0 minutes");
                    return;
                }

                ClearError(remindMinutesBeforeProperty);
                _remindMinutesBefore = value;
                RaisePropertyChanged(remindMinutesBeforeProperty);
            }
        }

        public BusyStatus ShowTimeAs
        {
            get { return _showTimeAs; }
            set
            {
                _showTimeAs = value;
                RaisePropertyChanged("ShowTimeAs");
            }
        }

        public RsvpStatus DownloadTypes
        {
            get { return _downloadTypes; }
            set
            {
                if (value == RsvpStatus.None)
                {
                    SetError("DownloadTypes", "You should have one or more RSVP statuses checked");
                    return;
                }
                _downloadTypes = value;
                RaisePropertyChanged("DownloadTypes");
            }
        }

        public string Category
        {
            get { return _category; }
            set
            {
                _category = value;
                RaisePropertyChanged("Category");
            }
        }
    }
}
