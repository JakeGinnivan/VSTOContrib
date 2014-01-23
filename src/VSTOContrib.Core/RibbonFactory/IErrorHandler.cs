using System;

namespace VSTOContrib.Core.RibbonFactory
{
    public interface IErrorHandler
    {
        bool Handle(Exception exception);
    }
}