using System;
using System.Collections.Generic;
using System.Linq;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;

namespace VSTOContrib.Core.Domain
{
    /// <summary>
    /// Represents an Office Context, i.e Document, Workbook, MailItem
    /// </summary>
    class OfficeContextDomain : IOfficeContextDomainWriter
    {
        readonly List<OfficeViewDomain> views = new List<OfficeViewDomain>();
        readonly ICheckWritable officeApplicationDomain;

        public OfficeContextDomain(object officeContext, OfficeApplicationDomain officeApplicationDomain)
        {
            this.officeApplicationDomain = officeApplicationDomain;
            officeApplicationDomain.NewView += OfficeApplicationDomainOnNewView;
            officeApplicationDomain.ViewClosed += OfficeApplicationDomainOnViewClosed;
            OfficeContext = officeContext;
        }

        void OfficeApplicationDomainOnNewView(OfficeViewDomain officeViewDomain)
        {
            if (officeViewDomain.Contexts.Any(v => Equals(v.OfficeContext, OfficeContext)))
                NewView(officeViewDomain);
        }

        void OfficeApplicationDomainOnViewClosed(OfficeViewDomain officeViewDomain)
        {
            if (officeViewDomain.Contexts.Any(v => Equals(v.OfficeContext, OfficeContext)))
                ViewClosed(officeViewDomain);
        }

        public event Action<OfficeViewDomain> NewView = _ => { };


        public event Action<OfficeViewDomain> ViewClosed = _ => { };

        public IEnumerable<OfficeViewDomain> Views { get { return views.ToArray(); } }

        public OfficeViewDomain ActiveView { get; private set; }

        public object OfficeContext { get; private set; }

        public IRibbonViewModel ViewModel { get; set; }

        void IOfficeContextDomainWriter.AddView(OfficeViewDomain view)
        {
            officeApplicationDomain.AssertWritable();
            views.Add(view);
        }
    }
}