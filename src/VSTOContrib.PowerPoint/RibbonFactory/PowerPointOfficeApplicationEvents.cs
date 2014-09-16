using System;
using Microsoft.Office.Interop.PowerPoint;
using VSTOContrib.Core;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.PowerPoint.RibbonFactory
{
    public class PowerPointOfficeApplicationEvents : IOfficeApplicationEvents
    {
        private Application powerPointApplication;
        const string CaptionSuffix = " - PowerPoint";
        const string PowerpointLpClassName = "PPTFrameClass";

        void PowerPointApplicationWindowActivate(Presentation pres, DocumentWindow window)
        {
            NewView(new NewViewEventArgs(ToOfficeWindow(window), pres, PowerPointRibbonType.PowerPointPresentation.GetEnumDescription()));
        }

        void PowerPointViewProviderNewPresentation(Presentation pres)
        {
            using(var documentWindows = pres.Windows.WithComCleanup())
            foreach (var documentWindow in documentWindows.Resource)
            {
                var powerPointPresentation = PowerPointRibbonType.PowerPointPresentation.GetEnumDescription();
                NewView(new NewViewEventArgs(ToOfficeWindow(documentWindow), pres, powerPointPresentation));
            }
        }

        public void Initialise(object application)
        {
            powerPointApplication = (Application) application;
            ((EApplication_Event)powerPointApplication).NewPresentation += PowerPointViewProviderNewPresentation;
            powerPointApplication.WindowActivate += PowerPointApplicationWindowActivate;
        }

        public event Action<NewViewEventArgs> NewView = _ => { };
        public event Action<OfficeWin32Window> ViewClosed = _ => { };
        public event Action<object> ContextClosed = _ => { };

        public OfficeWin32Window ToOfficeWindow(object view)
        {
            return new OfficeWin32Window(view, PowerpointLpClassName, CaptionSuffix);
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
        }

        /// <summary>
        /// Registers the open powerpoint presentations.
        /// </summary>
        public void RegisterOpenDocuments()
        {
            using (var presentations = powerPointApplication.Presentations.WithComCleanup())
            {
                foreach (Presentation presentation in presentations.Resource)
                {
                    using (var windows = presentation.Windows.WithComCleanup())
                    {
                        foreach (DocumentWindow window in windows.Resource)
                        {
                            PowerPointApplicationWindowActivate(presentation, window);
                        }
                    }
                }
            }
        }
    }
}