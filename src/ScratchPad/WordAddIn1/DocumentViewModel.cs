using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Word.RibbonFactory;

namespace TwitterResultsWordAddin
{
    [WordRibbonViewModel]
    public class DocumentViewModel : OfficeViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        private bool panelShown;
        private Document document;
        private ICustomTaskPaneWrapper myAddinTaskPane;
        bool ribbonVisible;

        public void Initialised(object context)
        {
            document = context as Document;
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
                    Child = new MyAddinPanel //This is a WPF User control
                    {
                        DataContext = new MyAddinPanelViewModel(GetHashCode()) //Viewmodel for the user control
                    }
                }, "Twitter Results");
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
