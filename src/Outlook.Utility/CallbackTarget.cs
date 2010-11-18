using System;

namespace Outlook.Utility
{
    internal class CallbackTarget 
    {
        private readonly Type _viewModelType;
        private readonly string _method;

        public CallbackTarget(Type viewModelType, string method)
        {
            _viewModelType = viewModelType;
            _method = method;
        }

        public string Method
        {
            get { return _method; }
        }

        public Type ViewModelType
        {
            get { return _viewModelType; }
        }
    }
}