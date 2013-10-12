using System.Drawing;
using GitHubForOutlook.Core.Features.CreateIssue;
using GitHubForOutlook.Core.Features.Settings;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using VSTOContrib.Core;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Outlook.RibbonFactory;

namespace GitHubForOutlook.Core.Ribbons
{
    [RibbonViewModel(OutlookRibbonType.OutlookExplorer)]
    public class GithubExplorerRibbon : OfficeViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        readonly ISettingsViewModel settingsViewModel;
        readonly ICreateIssueViewModel createIssuesViewModel;
        MailItem selectedMailItem;
        Explorer explorer;

        ICustomTaskPaneWrapper settingsTaskPane;
        ICustomTaskPaneWrapper createIssueTaskPane;

        public GithubExplorerRibbon(ISettingsViewModel settingsViewModel, ICreateIssueViewModel createIssuesViewModel)
        {
            this.settingsViewModel = settingsViewModel;
            this.createIssuesViewModel = createIssuesViewModel;
        }

        public Factory VstoFactory { get; set; }
        public bool MailItemSelected { get; set; }
        public IRibbonUI RibbonUi { get; set; }

        public void Initialised(object context)
        {
        }

        public void CreateIssue(IRibbonControl ribbonControl)
        {
            if (selectedMailItem == null) return;

            // If we are not authenticated, authenticate before allowing user to create issue
            if (string.IsNullOrEmpty(Properties.Settings.Default.AuthToken))
            {
                settingsViewModel.LoginCallback(() => CreateIssue(ribbonControl));
                settingsTaskPane.Visible = true;
                return;
            }

            if (settingsTaskPane.Visible)
                settingsTaskPane.Visible = false;

            createIssuesViewModel.CreateIssueFor(selectedMailItem);
            selectedMailItem = null;
            MailItemSelected = false;
            createIssueTaskPane.Visible = true;
        }

        public void ShowSettings(IRibbonControl ribbonControl)
        {
            if (createIssueTaskPane.Visible)
                createIssueTaskPane.Visible = false;
            settingsTaskPane.Visible = true;            
        }

        public void CurrentViewChanged(object currentView)
        {
            explorer = (Explorer)currentView;
            explorer.SelectionChange += ExplorerOnSelectionChange;
        }

        private void ExplorerOnSelectionChange()
        {
            using (var selection = explorer.Selection.WithComCleanup())
            {
                if (selection.Resource.Count == 1)
                {
                    object item = null;
                    MailItem mailItem = null;
                    try
                    {
                        item = selection.Resource[1];
                        mailItem = item as MailItem;
                        if (mailItem != null)
                        {
                            if (selectedMailItem != null)
                                selectedMailItem.ReleaseComObject();
                            selectedMailItem = mailItem;
                            MailItemSelected = true;
                        }
                        else
                        {
                            MailItemSelected = false;
                        }
                    }
                    finally
                    {
                        if (mailItem == null)
                            item.ReleaseComObject();
                    }
                }
                else
                {
                    MailItemSelected = false;
                }
            }
        }

        public Bitmap GetImage(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "createTask":
                    {
                        return Properties.Resources.gtfo32x32;
                    }
            }

            return null;
        }

        public void Cleanup()
        {
            explorer = null;
        }

        public void RegisterTaskPanes(Register register)
        {
            settingsTaskPane = register(() => new WpfPanelHost
            {
                Child = new SettingsControl
                {
                    DataContext = settingsViewModel
                }
            }, "GitHub Settings", initallyVisible:false);
            settingsViewModel.Init(settingsTaskPane);

            createIssueTaskPane = register(() => new WpfPanelHost
            {
                Child = new CreateIssueControl
                {
                    DataContext = createIssuesViewModel
                }
            }, "Create Issue", initallyVisible: false);
            createIssuesViewModel.Init(createIssueTaskPane);
        }
    }
}
