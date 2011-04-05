using System;

namespace VSTOContrib.Core.Tests.RibbonFactory.TestStubs
{
    [Flags]
    public enum TestRibbonTypes
    {
        RibbonType1 = 1,
        RibbonType2 = 1 << 1,
        RibbonType3 = 1 << 2
    }
}