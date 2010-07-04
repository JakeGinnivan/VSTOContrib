namespace Office.Utility
{
    internal class CallbackTarget 
    {
        private readonly IRibbonViewModel _viewModel;
        private readonly string _method;

        public CallbackTarget(IRibbonViewModel viewModel, string method)
        {
            _viewModel = viewModel;
            _method = method;
        }

        public string Method
        {
            get { return _method; }
        }

        public IRibbonViewModel ViewModel
        {
            get { return _viewModel; }
        }
    }
}