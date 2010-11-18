using System;
using FacebookToOutlook.Data.Adapters;
using FacebookToOutlook.Presentation.ViewModels;
using FacebookToOutlook.Presentation.Views;
using FacebookToOutlook.Properties;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Practices.ServiceLocation;
using Office.Utility;

namespace FacebookToOutlookAddin
{
    public partial class AppointmentInspectorRibbon
    {
        private Inspector _inspector;
        private AppointmentItem _appointment;
        private FacebookEventAdapter _appointmentAdapter;
        private CustomTaskPane _timecardTaskPane;
        private WpfPanelHost _control;

        private void AppointmentRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            _inspector = Context as Inspector;
            if (_inspector == null) return;
            ((InspectorEvents_Event)_inspector).Close += InspectorClose;
            _appointment = _inspector.CurrentItem as AppointmentItem;

            _appointmentAdapter = new FacebookEventAdapter(_appointment);
            if (!_appointmentAdapter.IsFacebookEvent)
            {
                facebookGroup.Visible = false;
                return;
            }

            CreateFacebookPanel();
        }

        private void CreateFacebookPanel()
        {
            _control = new WpfPanelHost
                              {
                                  Child = new FacebookEventPaneView
                                              {
                                                  DataContext = new FacebookEventPaneViewModel(_appointmentAdapter)
                                              }
                              };
            _timecardTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(_control, "Facebook", _inspector);

            _timecardTaskPane.Width = ServiceLocator.Current.GetInstance<IApplicationSettings>().AppointmentTaskPaneWidth;
            _timecardTaskPane.Visible = true;
            showFacebookPaneButton.Checked = true;
            _timecardTaskPane.VisibleChanged += TimecardTaskPaneVisibleChanged;
            TimecardTaskPaneVisibleChanged(this, EventArgs.Empty);
        }

        private void InspectorClose()
        {
            ((InspectorEvents_Event)_inspector).Close -= InspectorClose;
            _inspector = null;

            if (_timecardTaskPane == null) return;

            var appSettings = ServiceLocator.Current.GetInstance<IApplicationSettings>();
            appSettings.AppointmentTaskPaneWidth = _timecardTaskPane.Width;
            appSettings.Save();
            Globals.ThisAddIn.CustomTaskPanes.Remove(_timecardTaskPane);
            _control.Dispose();
        }

        private void TimecardTaskPaneVisibleChanged(object sender, EventArgs e)
        {
            showFacebookPaneButton.Checked = _timecardTaskPane.Visible;
        }

        private void TogglePanelVisibility(object sender, RibbonControlEventArgs e)
        {
            _timecardTaskPane.Visible = ((RibbonToggleButton)sender).Checked;
        }
    }
}
