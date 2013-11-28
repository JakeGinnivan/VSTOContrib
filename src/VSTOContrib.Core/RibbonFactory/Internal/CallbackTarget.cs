using System;

namespace VSTOContrib.Core.RibbonFactory.Internal
{
    internal class CallbackTarget
    {
        private readonly string ribbonType;
        private readonly string method;

        public CallbackTarget(string ribbonType, string method)
        {

            this.ribbonType = ribbonType;
            this.method = method;
        }

        public string Method
        {
            get { return method; }
        }

        public string RibbonType
        {
            get { return ribbonType; }
        }
    }
}