using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Outlook;

namespace OutlookQuickStart
{
    [FormRegionMessageClass(FormRegionMessageClassAttribute.Note)]
    [FormRegionName("OutlookQuickStart.FormRegion1")]
    public class FormRegistration : FormRegionBootstrapper
    {
    }
    
    public class FormRegionBootstrapper : IFormRegionFactory
    {
        public bool IsDisplayedForItem(object outlookItem, OlFormRegionMode formRegionMode, OlFormRegionSize formRegionSize)
        {
            return true;
        }

        public byte[] GetFormRegionStorage(object outlookItem, OlFormRegionMode formRegionMode, OlFormRegionSize formRegionSize)
        {
            return null;
        }

        public IFormRegion CreateFormRegion(FormRegion formRegion)
        {
            return null;
        }

        public FormRegionKindConstants Kind { get; private set; }
        public FormRegionManifest Manifest { get; private set; }
    }
}