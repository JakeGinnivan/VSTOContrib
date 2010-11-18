using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Outlook.Utility.RibbonFactory;

namespace RibbonFactoryDemo.Ribbons
{
    [RibbonViewModel(RibbonType.OutlookMailRead | RibbonType.OutlookMailCompose)]
    public class MailRibbon : IRibbonViewModel
    {
        public IRibbonUI RibbonUi { get; set; }

        public void Displayed(object context)
        {
            MessageBox.Show(((MailItem) ((Inspector)context).CurrentItem).Subject);
        }

        public void Cleanup()
        {
            
        }
    }
}
