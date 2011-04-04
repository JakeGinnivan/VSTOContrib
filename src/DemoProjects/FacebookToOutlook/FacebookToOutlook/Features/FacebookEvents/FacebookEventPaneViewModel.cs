using System;
using System.Diagnostics;
using System.Windows.Input;
using FacebookToOutlook.Core;
using FacebookToOutlook.Data.Adapters;
using FacebookToOutlook.Presentation.Commands;
using FacebookToOutlook.Properties;
using GalaSoft.MvvmLight;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Office.Contrib;
using Office.Contrib.RibbonFactory;
using Office.Contrib.RibbonFactory.Interfaces;
using Office.Outlook.Contrib.RibbonFactory;
using CustomTaskPane = Microsoft.Office.Tools.CustomTaskPane;

namespace FacebookToOutlook.Features.FacebookEvents
{
    [RibbonViewModel(OutlookRibbonType.OutlookAppointment)]
    public class FacebookEventPaneViewModel : ViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        private readonly IApplicationSettings _applicationSettings;
        private DelegateCommand _viewEventCommand;
        private FacebookEventAdapter _eventAdapter;
        private bool _isVisible;
        private WpfPanelHost _control;
        private CustomTaskPane _facebookTaskPane;
        private bool _isPanelShown;

        public FacebookEventPaneViewModel(IApplicationSettings applicationSettings)
        {
            _applicationSettings = applicationSettings;
        }

        public void Initialised(object context)
        {
            var inspector = ((Inspector) context);
            _eventAdapter = new FacebookEventAdapter((_AppointmentItem) inspector.CurrentItem);
            if (!_eventAdapter.IsFacebookEvent)
                SetIsRibbonVisible(false);
        }

        public void CurrentViewChanged(object currentView)
        {
            
        }

        public void RegisterTaskPanes(Register register)
        {
            _control = new WpfPanelHost
                              {
                                  Child = new FacebookEventPaneView
                                              {
                                                  DataContext = this
                                              }
                              };
            _facebookTaskPane = register(_control, "Facebook");
            _facebookTaskPane.Width = _applicationSettings.AppointmentTaskPaneWidth;
            _facebookTaskPane.Visible = true;
            SetPanelShown(true);
            _facebookTaskPane.VisibleChanged += FacebookTaskPaneVisibleChanged;
            FacebookTaskPaneVisibleChanged(this, EventArgs.Empty);
        }

        public override void Cleanup()
        {
            if (_facebookTaskPane == null) return;

            _applicationSettings.AppointmentTaskPaneWidth = _facebookTaskPane.Width;
            _applicationSettings.Save();
            _control.Dispose();
            base.Cleanup();
        }

        private void FacebookTaskPaneVisibleChanged(object sender, EventArgs e)
        {
            SetPanelShown(_facebookTaskPane.Visible);
        }

        public bool IsPanelShown(IRibbonControl control)
        {
            return _isPanelShown;
        }

        private void SetPanelShown(bool shown)
        {
            _isPanelShown = shown;
            if (RibbonUi != null)
                RibbonUi.Invalidate();
        }

        public bool IsRibbonVisible(IRibbonControl control)
        {
            return _isVisible;
        }

        private void SetIsRibbonVisible(bool visible)
        {
            _isVisible = visible;
            if (RibbonUi != null)
                RibbonUi.Invalidate();
        }

        public void TogglePanelVisibility(IRibbonControl control, bool pressed)
        {
            _isVisible = pressed;
            _facebookTaskPane.Visible = pressed;
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

        public IRibbonUI RibbonUi { get; set; }
    }
}
