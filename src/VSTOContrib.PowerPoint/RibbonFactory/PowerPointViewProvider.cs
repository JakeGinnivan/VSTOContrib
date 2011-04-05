using System;
using Microsoft.Office.Interop.PowerPoint;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.PowerPoint.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    public class PowerPointViewProvider : IViewProvider<PowerPointRibbonType>
    {
        private readonly Application _powerPointApplication;

        public PowerPointViewProvider(Application powerPointApplication)
        {
            _powerPointApplication = powerPointApplication;
            ((EApplication_Event)_powerPointApplication).NewPresentation += PowerPointViewProviderNewPresentation;
            _powerPointApplication.WindowActivate += PowerPointApplicationWindowActivate;
        }

        void PowerPointApplicationWindowActivate(Presentation pres, DocumentWindow window)
        {
            var handler = NewView;
            if (handler == null) return;

            handler(this, new NewViewEventArgs<PowerPointRibbonType>(
                              window, pres, PowerPointRibbonType.PowerPointPresentation));
        }

        void PowerPointViewProviderNewPresentation(Presentation pres)
        {
            var handler = NewView;
            if (handler == null) return;

            using(var documentWindows = pres.Windows.WithComCleanup())
            foreach (var documentWindow in documentWindows.Resource)
            {
                handler(this, new NewViewEventArgs<PowerPointRibbonType>(
                    documentWindow, pres,
                    PowerPointRibbonType.PowerPointPresentation));
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
        public event EventHandler<NewViewEventArgs<PowerPointRibbonType>> NewView;

        /// <summary>
        /// Occurs when [view closed].
        /// </summary>
        public event EventHandler<ViewClosedEventArgs> ViewClosed;

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
            using (var presentations = _powerPointApplication.Presentations.WithComCleanup())
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