namespace VSTOContrib.Core.RibbonFactory.Internal
{
    class NullContext
    {
        static NullContext()
        {
            Instance= new object();
        }

        public static object Instance { get; private set; }
    }
}