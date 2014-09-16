namespace VSTOContrib.Core.RibbonFactory
{
    public class NewViewEventArgs
    {
        readonly OfficeWin32Window viewInstance;
        readonly object viewContext;
        readonly string ribbonType;

        public NewViewEventArgs(OfficeWin32Window viewInstance, object viewContext, string ribbonType)
        {
            this.viewInstance = viewInstance;
            this.viewContext = viewContext;
            this.ribbonType = ribbonType;
        }

        public string RibbonType
        {
            get { return ribbonType; }
        }

        public OfficeWin32Window ViewInstance
        {
            get { return viewInstance; }
        }

        public object ViewContext
        {
            get { return viewContext; }
        }
    }
}