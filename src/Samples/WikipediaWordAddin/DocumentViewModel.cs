using System;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Word.RibbonFactory;
using Document = Microsoft.Office.Interop.Word.Document;

namespace WikipediaWordAddin
{
    [WordRibbonViewModel]
    public class DocumentViewModel : OfficeViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        readonly WikipediaResultsViewModel wikipediaResultsViewModel;
        bool panelShown, ribbonVisible;
        Document document;
        ICustomTaskPaneWrapper myAddinTaskPane;
        Microsoft.Office.Tools.Word.Document vstoDocument;

        public DocumentViewModel()
        {
            wikipediaResultsViewModel = new WikipediaResultsViewModel();
        }

        public void Initialised(object context)
        {
            document = context as Document;

            if (document != null)
            {
                vstoDocument = Globals.Factory.GetVstoObject(document);
                vstoDocument.SelectionChange += VstoDocumentOnSelectionChange;
            }
        }

        void VstoDocumentOnSelectionChange(object sender, SelectionEventArgs e)
        {
            using (e.WithComCleanup())
            using (var selection = e.Selection.WithComCleanup())
            {
                wikipediaResultsViewModel.Search(selection.Resource.Text);
            }
        }

        public bool RibbonVisible
        {
            get { return ribbonVisible; }
            set
            {
                ribbonVisible = value;
                RaisePropertyChanged(()=>RibbonVisible);
            }
        }

        public void CurrentViewChanged(object currentView)
        {
            RibbonVisible = document != null;
            panelShown = document != null;
        }

        public IRibbonUI RibbonUi { get; set; }

        public bool PanelShown
        {
            get { return panelShown; }
            set
            {
                if (panelShown == value) return;
                panelShown = value;
                myAddinTaskPane.Visible = value;
                RaisePropertyChanged(() => PanelShown);
            }
        }

        public void RegisterTaskPanes(Register register)
        {
            myAddinTaskPane = register(
                () => new WpfPanelHost
                {
                    Child = new WikipediaResultsView //This is a WPF User control
                    {
                        DataContext = wikipediaResultsViewModel //Viewmodel for the user control
                    }
                }, "Wikipedia Results");
            myAddinTaskPane.Visible = true;
            myAddinTaskPane.VisibleChanged += TaskPaneVisibleChanged;
            TaskPaneVisibleChanged(this, EventArgs.Empty);
        }

        public void Cleanup()
        {
            myAddinTaskPane.VisibleChanged -= TaskPaneVisibleChanged;
        }

        private void TaskPaneVisibleChanged(object sender, EventArgs e)
        {
            panelShown = myAddinTaskPane.Visible;
            RaisePropertyChanged(() => PanelShown);
        }
    }
}
