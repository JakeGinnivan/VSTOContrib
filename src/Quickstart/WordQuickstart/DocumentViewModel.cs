using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Word.RibbonFactory;

namespace WordQuickstart
{
    [WordRibbonViewModel]
    public class DocumentViewModel : OfficeViewModelBase, IRibbonViewModel
    {
        Document document;
        Window currentWindow;

        public IRibbonUI RibbonUi { get; set; }

        public void Initialised(object context)
        {
            document = (Document) context;
        }

        public void CurrentViewChanged(object currentView)
        {
            // This is for a single document being opened in multiple windows
            currentWindow = (Window) currentView;
        }

        public void Cleanup()
        {

        }
    }
}