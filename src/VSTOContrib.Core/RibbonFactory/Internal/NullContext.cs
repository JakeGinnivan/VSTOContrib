namespace VSTOContrib.Core.RibbonFactory.Internal
{
    class NullContext
    {
        static NullContext()
        {
            Instance = new NullContext();
        }

        public static object Instance { get; private set; }
    }
}