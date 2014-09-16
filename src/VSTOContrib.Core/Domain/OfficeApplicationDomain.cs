using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;

namespace VSTOContrib.Core.Domain
{
    class OfficeApplicationDomain : ICheckWritable
    {
        readonly IOfficeApplicationEvents officeApplicationEvents;
        readonly object nullContext = new object();
        readonly Dictionary<object, OfficeContextDomain> contexts = new Dictionary<object, OfficeContextDomain>();
        readonly Dictionary<OfficeWin32Window, OfficeViewDomain> views = new Dictionary<OfficeWin32Window, OfficeViewDomain>();
        bool isWriting;

        public OfficeApplicationDomain(IOfficeApplicationEvents officeApplicationEvents)
        {
            this.officeApplicationEvents = officeApplicationEvents;
            officeApplicationEvents.NewView += OfficeApplicationEventsOnNewView;
            officeApplicationEvents.ViewClosed += OfficeApplicationEventsOnViewClosed;
            officeApplicationEvents.ContextClosed += OfficeApplicationEventsOnContextClosed;
        }

        void OfficeApplicationEventsOnNewView(NewViewEventArgs args)
        {
            if (!contexts.ContainsKey(args.ViewContext ?? nullContext) && !views.ContainsKey(args.ViewInstance))
            {
                var newContext = new OfficeContextDomain(args.ViewContext, this);
                var newView = new OfficeViewDomain(args.ViewInstance, this);

                AddView(newContext, newView);
                AddContext(newContext, newView);
                NewView(newView);
                NewContext(newContext);
            }
            else if (!contexts.ContainsKey(args.ViewContext ?? nullContext))
            {
                var newContext = new OfficeContextDomain(args.ViewContext, this);
                var existingView = views[args.ViewInstance];
                AddContext(newContext, existingView);
                NewContext(newContext);
            }
            else if (!views.ContainsKey(args.ViewInstance))
            {
                var existingContext = contexts[args.ViewContext ?? nullContext];
                var newView = new OfficeViewDomain(args.ViewInstance, this);
                AddView(existingContext, newView);
                NewView(newView);
            }
        }

        void OfficeApplicationEventsOnViewClosed(OfficeWin32Window args)
        {

        }

        void OfficeApplicationEventsOnContextClosed(object officeContext)
        {

        }

        public event Action<OfficeViewDomain> NewView = _ => { };

        public event Action<OfficeContextDomain> NewContext = _ => { };

        public event Action<OfficeViewDomain> ViewClosed = _ => { };

        public event Action<OfficeContextDomain> ContextClosed = _ => { };

        public void RibbonLoaded(IRibbonUI ribbonUi)
        {

        }

        public OfficeViewDomain GetView(OfficeWin32Window window)
        {
            return null;
        }

        public OfficeContextDomain GetContext(object officeContext)
        {
            if (officeContext == null)
            {
                if (!contexts.ContainsKey(nullContext))
                {
                    OfficeApplicationEventsOnNewView(new NewViewEventArgs(officeApplicationEvents.ActiveWindow, ));
                }
                return contexts[nullContext];
            }

            if (officeContext is OfficeContextDomain)
                throw new ArgumentException("Context should be an office context", "officeContext");

            if (contexts.ContainsKey(officeContext))
                return contexts[officeContext];

            return null;
        }

        public void AssertWritable()
        {
            if (!isWriting)
                throw new InvalidOperationException("Domain write methods are not supposed to be called by anything other than the domain");
        }

        void AddContext(OfficeContextDomain newContext, OfficeViewDomain newView)
        {
            isWriting = true;
            contexts.Add(newContext.OfficeContext, newContext);
            ((IOfficeViewDomainWriter)newView).AddContext(newContext);
            isWriting = false;
        }

        void AddView(OfficeContextDomain newContext, OfficeViewDomain newView)
        {
            isWriting = true;
            views.Add(newView.Window, newView);
            ((IOfficeContextDomainWriter)newContext).AddView(newView);
            isWriting = false;
        }
    }
}