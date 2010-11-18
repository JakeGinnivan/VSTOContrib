using Microsoft.Office.Core;
using Office.Utility;

namespace FacebookToOutlookAddin.Ribbons
{
    public class AppointmentRibbon : IRibbonViewModel
    {
        public RibbonType Type
        {
            get { return RibbonType.OutlookAppointment; }
        }

        public void TogglePanelVisibility(IRibbonControl control, bool pressed)
        {
            
        }

        public IRibbonUI RibbonUi
        {
            get;
            set;
        }
    }
}
