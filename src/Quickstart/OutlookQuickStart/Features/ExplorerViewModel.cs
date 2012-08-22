using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Outlook.RibbonFactory;

namespace OutlookQuickStart.Features
{
    [OutlookRibbonViewModel(OutlookRibbonType.OutlookExplorer)]
    public class ExplorerViewModel : OfficeViewModelBase, IRibbonViewModel
    {
        Folder context;
        Explorer explorer;

        public IRibbonUI RibbonUi { get; set; }
        public void Initialised(object context)
        {
            this.context = (Folder) context;
        }

        public void EmailItems(IRibbonControl control)
        {
            if (explorer.Selection.Count == 1)
            {
                var title = ((MailItem) explorer.Selection[1]).Subject;
            }
        }

        public void CurrentViewChanged(object currentView)
        {
            explorer = (Explorer) currentView;
        }

        public void Cleanup()
        {
        }
    }
}