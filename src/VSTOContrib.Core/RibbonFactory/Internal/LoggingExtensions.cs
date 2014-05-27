namespace VSTOContrib.Core.RibbonFactory.Internal
{
    static class LoggingExtensions
    {
        public static string ToLogFormat(this object thing)
        {
            if (thing == null) 
                return "null";
            return string.Format("{0} ({1})", thing.GetType().Name, thing.GetHashCode());
        }
    }
}