using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Core;
using VSTOContrib.Core.RibbonFactory.Internal;

namespace VSTOContrib.Core.Domain
{
    /// <summary>
    /// Represents an Office View/Window
    /// </summary>
    class OfficeViewDomain : IOfficeViewDomainWriter
    {
        readonly List<OfficeContextDomain> contexts = new List<OfficeContextDomain>();
        readonly ICheckWritable officeApplicationDomain;

        public OfficeViewDomain(OfficeWin32Window viewInstance, OfficeApplicationDomain officeApplicationDomain)
        {
            this.officeApplicationDomain = officeApplicationDomain;
            officeApplicationDomain.NewContext += OfficeApplicationDomainOnNewContext;
            officeApplicationDomain.ContextClosed += OfficeApplicationDomainOnContextClosed;
            Window = viewInstance;
        }

        void OfficeApplicationDomainOnNewContext(OfficeContextDomain officeContextDomain)
        {
            if (officeContextDomain.Views.Any(v => Equals(v.Window, Window)))
                NewContext(officeContextDomain);
        }

        void OfficeApplicationDomainOnContextClosed(OfficeContextDomain officeContextDomain)
        {
            if (officeContextDomain.Views.Any(v => Equals(v.Window, Window)))
                ContextClosed(officeContextDomain);
        }

        public event Action<OfficeContextDomain> NewContext = _ => { };


        public event Action<OfficeContextDomain> ContextClosed = _ => { };

        public IEnumerable<OfficeContextDomain> Contexts { get { return contexts.ToArray(); } }

        public OfficeWin32Window Window { get; private set; }

        public IRibbonUI RibbonUI { get; private set; }

        void IOfficeViewDomainWriter.AddContext(OfficeContextDomain context)
        {
            officeApplicationDomain.AssertWritable();
            contexts.Add(context);
        }
    }
}