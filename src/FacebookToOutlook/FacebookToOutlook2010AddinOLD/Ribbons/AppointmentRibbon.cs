using System;
using Microsoft.Office.Core;
using Office.Utility;
using Outlook.Utility.RibbonFactory;

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

        public void Displayed(object context)
        {
            throw new NotImplementedException();
        }

        public void Cleanup()
        {
            throw new NotImplementedException();
        }
    }
}
