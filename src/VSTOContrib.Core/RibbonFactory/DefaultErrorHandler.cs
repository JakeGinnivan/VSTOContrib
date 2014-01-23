using System;
using System.Diagnostics;

namespace VSTOContrib.Core.RibbonFactory
{
    class DefaultErrorHandler : IErrorHandler
    {
        public bool Handle(Exception exception)
        {
            Trace.TraceError(exception.ToString());
            return false;
        }
    }
}