using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.PowerPoint.RibbonFactory;

namespace AddTextAddin
{
    [PowerPointRibbonViewModel]
    public class PresentationViewModel : OfficeViewModelBase, IRibbonViewModel
    {
        Microsoft.Office.Interop.PowerPoint.Presentation presentation;
        Microsoft.Office.Interop.PowerPoint.Application application;

        public IRibbonUI RibbonUi { get; set; }

        public Factory VstoFactory { get; set; }
        public object CurrentView { get; set; }

        public void Initialised(object context)
        {
            presentation = (Microsoft.Office.Interop.PowerPoint.Presentation) context;
            if (presentation == null) return;
            application = presentation.Application;
            application.PresentationNewSlide += ApplicationOnPresentationNewSlide;
        }

        void ApplicationOnPresentationNewSlide(Microsoft.Office.Interop.PowerPoint.Slide sld)
        {
            if (!AddTextEnabled) return;
            Microsoft.Office.Interop.PowerPoint.Shape textBox = sld.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            textBox.TextFrame.TextRange.InsertAfter("This text was added by using code.");
        }

        public bool AddTextEnabled { get; set; }

        public void Cleanup()
        {
        }
    }
}
