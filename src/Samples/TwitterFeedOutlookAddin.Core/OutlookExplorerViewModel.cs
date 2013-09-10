using System;
using System.Diagnostics;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Outlook.RibbonFactory;

namespace TwitterFeedOutlookAddin.Core
{
    [OutlookRibbonViewModel(OutlookRibbonType.OutlookExplorer)]
    public class OutlookExplorerViewModel : OfficeViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        private bool panelShown;
        private ICustomTaskPaneWrapper taskPane;

        public bool PanelShown
        {
            get { return panelShown; }
            set
            {
                if (panelShown == value) return;
                panelShown = taskPane.Visible = value;
                OnPropertyChanged(() => PanelShown);
            }
        }

        public void RegisterTaskPanes(Register register)
        {
            Debug.WriteLine("OutlookExplorerViewModel: RegisterTaskPanes " + GetHashCode());

            taskPane = register(() => new WpfPanelHost(), "Test Taskpane");
            taskPane.Visible = PanelShown = true;
            taskPane.VisibleChanged += TaskPaneVisibleChanged;
            TaskPaneVisibleChanged(this, EventArgs.Empty);
        }

        public IRibbonUI RibbonUi { get; set; }
        public Factory VstoFactory { get; set; }

        public void Cleanup()
        {
            Debug.WriteLine("OutlookExplorerViewModel: Cleanup " + GetHashCode());

            taskPane.VisibleChanged -= TaskPaneVisibleChanged;
        }

        /// <summary>
        ///     VSTOContrib provides currentView of Explorer or Inspector
        /// </summary>
        /// <param name="currentView">Explorer|Inspector</param>
        public void CurrentViewChanged(object currentView)
        {
            Debug.WriteLine("OutlookExplorerViewModel: CurrentViewChanged " + GetHashCode());
        }

        /// <summary>
        ///     VSTOContrib provides a context of Explorer.CurrentFolder or Inspector.CurrentItem
        /// </summary>
        /// <param name="context">Explorer.CurrentFolder|Inspector.CurrentItem</param>
        public void Initialised(object context)
        {
            Debug.WriteLine("OutlookExplorerViewModel: Initialised " + GetHashCode());
        }

        private void TaskPaneVisibleChanged(object sender, EventArgs e)
        {
            if (panelShown == taskPane.Visible) return;
            panelShown = taskPane.Visible;
            OnPropertyChanged(() => PanelShown);
        }
    }
}