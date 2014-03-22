using System;
using Microsoft.Office.Interop.PowerPoint;
using VSTOContrib.Core;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.PowerPoint.RibbonFactory
{
    /// <summary>
    /// PowerPoint View Provider
    /// </summary>
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
            var handler = NewView;
            if (handler == null) return;

            handler(this, new NewViewEventArgs(window, pres, PowerPointRibbonType.PowerPointPresentation.GetEnumDescription()));
        }

        void PowerPointViewProviderNewPresentation(Presentation pres)
        {
            var handler = NewView;
            if (handler == null) return;

            using(var documentWindows = pres.Windows.WithComCleanup())
            foreach (var documentWindow in documentWindows.Resource)
            {
                handler(this, new NewViewEventArgs(
                    documentWindow, pres,
                    PowerPointRibbonType.PowerPointPresentation.GetEnumDescription()));
            }
        }

        /// <summary>
        /// Initialises this instance.
        /// </summary>
        public void Initialise()
        {
        }

        /// <summary>
        /// Occurs when [new view].
        /// </summary>
        public event EventHandler<NewViewEventArgs> NewView;

        /// <summary>
        /// Occurs when [view closed].
        /// </summary>
        public event EventHandler<ViewClosedEventArgs> ViewClosed;

        public event EventHandler<HideCustomTaskPanesForContextEventArgs> UpdateCustomTaskPanesVisibilityForContext;

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