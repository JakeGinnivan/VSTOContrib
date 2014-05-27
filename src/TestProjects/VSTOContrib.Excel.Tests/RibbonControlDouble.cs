using Microsoft.Office.Core;

namespace VSTOContrib.Excel.Tests
{
    public class RibbonControlDouble : IRibbonControl
    {
        public RibbonControlDouble(string id, object context, string tag)
        {
            Context = context;
            Id = id;
            Tag = tag;
        }

        public string Id { get; private set; }
        public object Context { get; private set; }
        public string Tag { get; private set; }
    }
}