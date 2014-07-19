using System;
using Microsoft.Office.Interop.PowerPoint;
using VSTOContrib.Core;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.PowerPoint.RibbonFactory
{
    public class PowerPointViewProvider : IViewProvider
    {
        private Application powerPointApplication;
        const string CaptionSuffix = " - PowerPoint";
        const string PowerpointLpClassName = "PPTFrameClass";

        void PowerPointApplicationWindowActivate(Presentation pres, DocumentWindow window)
        {
            NewView(this, new NewViewEventArgs(new OfficeWin32Window(window, PowerpointLpClassName, CaptionSuffix), pres, PowerPointRibbonType.PowerPointPresentation.GetEnumDescription()));
        }

        void PowerPointViewProviderNewPresentation(Presentation pres)
        {
            using(var documentWindows = pres.Windows.WithComCleanup())
            foreach (var documentWindow in documentWindows.Resource)
            {
                var powerPointPresentation = PowerPointRibbonType.PowerPointPresentation.GetEnumDescription();
                NewView(this, new NewViewEventArgs(new OfficeWin32Window(documentWindow, PowerpointLpClassName, CaptionSuffix), pres, powerPointPresentation));
            }
        }

        public void Initialise(object application)
        {
            powerPointApplication = (Application) application;
            ((EApplication_Event)powerPointApplication).NewPresentation += PowerPointViewProviderNewPresentation;
            powerPointApplication.WindowActivate += PowerPointApplicationWindowActivate;
        }

        public event EventHandler<NewViewEventArgs> NewView = (sender, args) => { };
        public event EventHandler<ViewClosedEventArgs> ViewClosed = (sender, args) => { };

        /// <summary>
        /// Cleanups the references to a view.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <param name="context"></param>
        public void CleanupReferencesTo(OfficeWin32Window view, object context)
        {
            
        }

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