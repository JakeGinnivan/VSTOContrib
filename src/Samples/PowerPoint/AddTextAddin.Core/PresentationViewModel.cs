using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.PowerPoint.RibbonFactory;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace AddTextAddin.Core
{
    [PowerPointRibbonViewModel]
    public class PresentationViewModel : OfficeViewModelBase, IRibbonViewModel
    {
        PowerPoint.Presentation presentation;
        PowerPoint.Application application;

        public IRibbonUI RibbonUi { get; set; }

        public Factory VstoFactory { get; set; }
        public object CurrentView { get; set; }

        public void Initialised(object context)
        {
            presentation = (PowerPoint.Presentation) context;
            if (presentation == null) return;
            application = presentation.Application;
            application.PresentationNewSlide += ApplicationOnPresentationNewSlide;
        }

        void ApplicationOnPresentationNewSlide(PowerPoint.Slide sld)
        {
            if (!AddTextEnabled) return;
            PowerPoint.Shape textBox = sld.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            textBox.TextFrame.TextRange.InsertAfter("This text was added by using code.");
        }

        public bool AddTextEnabled { get; set; }

        public void Cleanup()
        {
        }
    }
}
