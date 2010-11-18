using Outlook.Utility.RibbonFactory;

namespace Outlook.Utility
{
    internal class CallbackTarget 
    {
        private readonly RibbonType _ribbonType;
        private readonly string _method;

        public CallbackTarget(RibbonType ribbonType, string method)
        {
            _ribbonType = ribbonType;
            _method = method;
        }

        public string Method
        {
            get { return _method; }
        }

        public RibbonType RibbonType
        {
            get { return _ribbonType; }
        }
    }
}