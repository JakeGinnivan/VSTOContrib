namespace VSTOContrib.Core.RibbonFactory.Internal
{
    internal class ViewModelKey
    {
        public object View { get; private set; }
        public object Context { get; private set; }

        public ViewModelKey(object view, object context)
        {
            View = view;
            Context = context;
        }
    }
}