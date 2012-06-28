using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Outlook.RibbonFactory;

namespace OutlookQuickStart.Features
{
    [OutlookRibbonViewModel(OutlookRibbonType.OutlookMailRead)]
    public class MailItemViewModel : OfficeViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        bool panelShown;
        MailItem mailItem;
        ICustomTaskPaneWrapper myAddinTaskPane;

        public void Initialised(object context)
        {
            mailItem = (MailItem) context;
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
                    //Child = new MyAddinPanel //This is a WPF User control
                    //{
                    //    DataContext = new MyAddinPanelViewModel(this) //Viewmodel for the user control
                    //}
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

        void TaskPaneVisibleChanged(object sender, EventArgs e)
        {
            panelShown = myAddinTaskPane.Visible;
            RaisePropertyChanged(() => PanelShown);
        }
    }
}