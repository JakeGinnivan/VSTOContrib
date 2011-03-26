using Microsoft.Office.Interop.PowerPoint;

namespace Office.PowerPoint.Contrib
{
    public class PresentationWrapper
    {
        public Presentation Presentation { get; set; }

        public PresentationWrapper(Presentation presentation)
        {
            Presentation = presentation;
            //No way to figure out if a presentation is closed?
        }
    }
}