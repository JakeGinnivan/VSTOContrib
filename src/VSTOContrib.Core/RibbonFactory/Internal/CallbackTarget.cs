using System;

namespace VSTOContrib.Core.RibbonFactory.Internal
{
    internal class CallbackTarget
    {
        private readonly string ribbonType;
        private readonly string method;
        private readonly string dynamicContext;

        public CallbackTarget(string ribbonType, string dynamicContext, string method)
        {

            this.ribbonType = ribbonType;
            this.dynamicContext = dynamicContext;
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

        public string DynamicContext
        {
            get { return dynamicContext; }
        }
    }
}