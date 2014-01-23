using System;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using VSTOContrib.Core;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Word.RibbonFactory;
using WikipediaWordAddin.Core.WpfControls;
using Document = Microsoft.Office.Interop.Word.Document;
using Factory = Microsoft.Office.Tools.Factory;

namespace WikipediaWordAddin.Core.OfficeContexts
{
    [WordRibbonViewModel]
    public class DocumentViewModel : OfficeViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        readonly WikipediaResultsViewModel wikipediaResultsViewModel;
        bool panelShown, ribbonVisible;
        Document document;
        ICustomTaskPaneWrapper myAddinTaskPane;
        Microsoft.Office.Tools.Word.Document vstoDocument;

        public DocumentViewModel(WikipediaResultsViewModel wikipediaResultsViewModel)
        {
            this.wikipediaResultsViewModel = wikipediaResultsViewModel;
        }

        public IRibbonUI RibbonUi { get; set; }

        public Factory VstoFactory { get; set; }

        public void Initialised(object context)
        {
            document = context as Document;

            if (document != null)
            {
                vstoDocument = ((ApplicationFactory)VstoFactory).GetVstoObject(document);
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
                OnPropertyChanged(()=>RibbonVisible);
            }
        }

        public void CurrentViewChanged(object currentView)
        {
            RibbonVisible = document != null;
        }
        
        public bool PanelShown
        {
            get { return panelShown; }
            set
            {
                if (panelShown == value) return;
                panelShown = value;
                myAddinTaskPane.Visible = value;
                OnPropertyChanged(() => PanelShown);
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
                }, "Wikipedia Results", initallyVisible: false);
            myAddinTaskPane.VisibleChanged += TaskPaneVisibleChanged;
        }

        public void Cleanup()
        {
            myAddinTaskPane.VisibleChanged -= TaskPaneVisibleChanged;
            vstoDocument.SelectionChange -= VstoDocumentOnSelectionChange;
        }

        private void TaskPaneVisibleChanged(object sender, EventArgs e)
        {
            panelShown = myAddinTaskPane.Visible;
            OnPropertyChanged(() => PanelShown);
        }
    }
}
