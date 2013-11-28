using System;

namespace VSTOContrib.Core.RibbonFactory.Internal
{
    internal class CallbackTarget<TRibbonType> where TRibbonType : struct 
    {
        private readonly TRibbonType _ribbonType;
        private readonly string _method;

        public CallbackTarget(TRibbonType ribbonType, string method)
        {
            if (!typeof(TRibbonType).IsEnum) throw new ArgumentException("Ribbon type must be enum");

            _ribbonType = ribbonType;
            _method = method;
        }

        public string Method
        {
            get { return _method; }
        }

        public TRibbonType RibbonType
        {
            get { return _ribbonType; }
        }
    }
}