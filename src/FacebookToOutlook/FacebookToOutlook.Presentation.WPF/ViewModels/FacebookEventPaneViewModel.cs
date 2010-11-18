using System.Diagnostics;
using System.Windows.Input;
using FacebookToOutlook.Core;
using FacebookToOutlook.Presentation.Commands;
using GalaSoft.MvvmLight;

namespace FacebookToOutlook.Presentation.ViewModels
{
    public class FacebookEventPaneViewModel : ViewModelBase
    {
        private readonly IFacebookEvent _eventAdapter;
        private DelegateCommand _viewEventCommand;

        public FacebookEventPaneViewModel(IFacebookEvent eventAdapter)
        {
            _eventAdapter = eventAdapter;
        }

        public bool Attending
        {
            get
            {
                return _eventAdapter.RsvpStatus  == RsvpStatus.Attending;
            }
        }

        public bool MaybeAttending
        {
            get
            {
                return _eventAdapter.RsvpStatus == RsvpStatus.Unsure;
            }
        }

        public bool NotAttending
        {
            get
            {
                return _eventAdapter.RsvpStatus == RsvpStatus.Declined;
            }
        }

        public ICommand ViewEventCommand
        {
            get
            {
                return _viewEventCommand ??
                    (_viewEventCommand = new DelegateCommand(ViewEvent));
            }
        }

        private void ViewEvent()
        {
            Process.Start(string.Format("http://www.facebook.com/event.php?eid={0}", _eventAdapter.EventId));
        }
    }
}
