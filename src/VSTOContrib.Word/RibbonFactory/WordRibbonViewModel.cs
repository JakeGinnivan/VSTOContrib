using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Word.RibbonFactory
{
    [WordRibbonViewModel]
    public abstract class WordRibbonViewModel : OfficeViewModelBase, IRibbonViewModel
    {
        /// <summary>
        /// The associated ribbon, you can invalidate controls getting them to refresh
        /// their state through the IRibbonUI.
        /// </summary>
        /// <value>The ribbon UI.</value>
        public IRibbonUI RibbonUi { get; set; }

        /// <summary>
        /// Used to get the VSTO versions of the context. For example
        /// 
        /// var vstodocument = ((Microsoft.Office.Tools.Word.ApplicationFactory)VstoFactory).GetVstoObject(document);
        /// </summary>
        public Factory VstoFactory { get; set; }

        object IRibbonViewModel.CurrentView { get; set; }

        public Window CurrentWindow
        {
            get { return ((IRibbonViewModel)this).CurrentView as Window; }
        }

        public ProtectedViewWindow CurrentProtectedWindow
        {
            get { return ((IRibbonViewModel)this).CurrentView as ProtectedViewWindow; }
        }

        public bool IsWindowProtected { get { return ((IRibbonViewModel)this).CurrentView is ProtectedViewWindow; } }

        /// <summary>
        /// Called when the window that the ribbon is shown in is opened
        /// </summary>
        /// <param name="context">The context.</param>
        void IRibbonViewModel.Initialised(object context)
        {
            Initialised(context as Document);
        }

        public abstract void Initialised(Document document);

        /// <summary>
        /// Cleanups this instance.
        /// </summary>
        public abstract void Cleanup();
    }
}
