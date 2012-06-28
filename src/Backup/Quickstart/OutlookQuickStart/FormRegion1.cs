using Microsoft.Office.Tools.Outlook;

namespace OutlookQuickStart
{
    partial class FormRegion1
    {
        [FormRegionMessageClass(FormRegionMessageClassAttribute.Note)]
        [FormRegionName("OutlookQuickStart.FormRegion1")]
        public partial class FormRegion1Factory
        {
            // Occurs before the form region is initialized.
            // To prevent the form region from appearing, set e.Cancel to true.
            // Use e.OutlookItem to get a reference to the current Outlook item.
            private void FormRegion1FactoryFormRegionInitializing(object sender, FormRegionInitializingEventArgs e)
            {
            }
        }

        // Occurs before the form region is displayed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void FormRegion1FormRegionShowing(object sender, System.EventArgs e)
        {
        }

        // Occurs when the form region is closed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void FormRegion1FormRegionClosed(object sender, System.EventArgs e)
        {
        }
    }
}
