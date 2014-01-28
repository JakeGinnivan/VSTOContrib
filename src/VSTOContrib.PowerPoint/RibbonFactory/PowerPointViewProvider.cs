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
        private readonly Application powerPointApplication;

        public PowerPointViewProvider(Application powerPointApplication)
        {
            this.powerPointApplication = powerPointApplication;
            ((EApplication_Event)this.powerPointApplication).NewPresentation += PowerPointViewProviderNewPresentation;
            this.powerPointApplication.WindowActivate += PowerPointApplicationWindowActivate;
        }

        void PowerPointApplicationWindowActivate(Presentation pres, DocumentWindow window)
        {
            NewView(this, new NewViewEventArgs(window, pres, PowerPointRibbonType.PowerPointPresentation.GetEnumDescription()));
        }

        void PowerPointViewProviderNewPresentation(Presentation pres)
        {
            using(var documentWindows = pres.Windows.WithComCleanup())
            foreach (var documentWindow in documentWindows.Resource)
            {
                NewView(this, new NewViewEventArgs(
                    documentWindow, pres,
                    PowerPointRibbonType.PowerPointPresentation.GetEnumDescription()));
            }
        }

        public void Initialise()
        {
        }

        public event EventHandler<NewViewEventArgs> NewView = (sender, args) => { };
        public event EventHandler<ViewClosedEventArgs> ViewClosed = (sender, args) => { };
        public event EventHandler<HideCustomTaskPanesForContextEventArgs> UpdateCustomTaskPanesVisibilityForContext
            = (sender, args) => { };

        /// <summary>
        /// Cleanups the references to a view.
        /// </summary>
        /// <param name="view">The view.</param>
        /// <param name="context"></param>
        public void CleanupReferencesTo(object view, object context)
        {
            
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