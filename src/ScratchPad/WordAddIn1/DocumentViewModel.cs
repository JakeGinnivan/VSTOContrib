using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Word.RibbonFactory;

namespace WordAddIn1
{
    [WordRibbonViewModel]
    public class DocumentViewModel : OfficeViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        private bool panelShown;
        private Document document;
        private ICustomTaskPaneWrapper myAddinTaskPane;

        public void Initialised(object context)
        {
            document = (Document)context;
        }

        public void CurrentViewChanged(object currentView)
        {
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
                        DataContext = new MyAddinPanelViewModel() //Viewmodel for the user control
                    }
                }, "MyAddin Awesome Taskpane");
            myAddinTaskPane.Visible = true;
            PanelShown = true;
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
