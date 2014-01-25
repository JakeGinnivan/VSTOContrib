using System;
using Microsoft.Office.Tools.Word;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Word.RibbonFactory;
using WikipediaWordAddin.Core.WpfControls;
using Document = Microsoft.Office.Interop.Word.Document;

namespace WikipediaWordAddin.Core.OfficeContexts
{
    public class DocumentViewModel : WordRibbonViewModel, IRegisterCustomTaskPane
    {
        readonly WikipediaResultsViewModel wikipediaResultsViewModel;
        bool panelShown, ribbonVisible;
        ICustomTaskPaneWrapper wikipediaResultsTaskPane;
        Microsoft.Office.Tools.Word.Document vstoDocument;

        public DocumentViewModel(WikipediaResultsViewModel wikipediaResultsViewModel)
        {
            this.wikipediaResultsViewModel = wikipediaResultsViewModel;
        }

        public override void Initialised(Document document)
        {
            if (document != null)
            {
                vstoDocument= ((ApplicationFactory)VstoFactory).GetVstoObject(document);
                vstoDocument.SelectionChange += VstoDocumentOnSelectionChange;
                RibbonVisible = true;
            }
        }

        void VstoDocumentOnSelectionChange(object sender, SelectionEventArgs e)
        {
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
                OnPropertyChanged(() => RibbonVisible);
            }
        }

        public bool PanelShown
        {
            get { return panelShown; }
            set
            {
                if (panelShown == value) return;
                panelShown = value;
                wikipediaResultsTaskPane.Visible = value;
                OnPropertyChanged(() => PanelShown);
            }
        }

        public void RegisterTaskPanes(Register register)
        {
            wikipediaResultsTaskPane = register(
                () => new WpfPanelHost
                {
                    Child = new WikipediaResultsView //This is a WPF User control
                    {
                        DataContext = wikipediaResultsViewModel //Viewmodel for the user control
                    }
                }, "Wikipedia Results");
            wikipediaResultsTaskPane.VisibleChanged += TaskPaneVisibleChanged;
        }

        public override void Cleanup()
        {
            wikipediaResultsTaskPane.VisibleChanged -= TaskPaneVisibleChanged;
            vstoDocument.SelectionChange -= VstoDocumentOnSelectionChange;
        }

        private void TaskPaneVisibleChanged(object sender, EventArgs e)
        {
            panelShown = wikipediaResultsTaskPane.Visible;
            OnPropertyChanged(() => PanelShown);
        }
    }
}
