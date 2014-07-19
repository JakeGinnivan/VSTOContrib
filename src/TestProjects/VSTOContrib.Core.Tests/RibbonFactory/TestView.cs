using System;
using System.Collections.Generic;

namespace VSTOContrib.Core.Tests.RibbonFactory
{
    public class TestView
    {
        static readonly Dictionary<object, IntPtr> Lookup = new Dictionary<object, IntPtr>();
        static int counter = 1;

        static TestView()
        {
            OfficeWin32Window.ResolveWindowHandle = o =>
            {
                if (!Lookup.ContainsKey(o))
                    Lookup.Add(o, new IntPtr(counter++));
                return Lookup[o];
            };
        }

        public TestWindowContext Context { get; set; }

        public OfficeWin32Window ToOfficeWin32Window()
        {
            return new OfficeWin32Window(this, string.Empty, string.Empty);
        }
    }
}