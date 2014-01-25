using System;
using System.Drawing;
using GitHubForOutlook.Core.Features;
using GitHubForOutlook.Core.Features.CreateIssue;
using GitHubForOutlook.Core.Features.Settings;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using VSTOContrib.Core;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Outlook.RibbonFactory;

namespace GitHubForOutlook.Core.Ribbons
{
    [RibbonViewModel(OutlookRibbonType.OutlookExplorer)]
    public class GithubExplorerRibbon : OfficeViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        readonly ITaskPaneContentHost contentHost;
        readonly ISettingsViewModel settingsViewModel;
        readonly Func<ICreateIssueViewModel> createIssuesViewModelFactory;
        MailItem selectedMailItem;
        Explorer explorer;

        public GithubExplorerRibbon(
            ITaskPaneContentHost contentHost, 
            ISettingsViewModel settingsViewModel, 
            Func<ICreateIssueViewModel> createIssuesViewModelFactory)
        {
            this.contentHost = contentHost;
            this.settingsViewModel = settingsViewModel;
            this.createIssuesViewModelFactory = createIssuesViewModelFactory;
        }

        public Factory VstoFactory { get; set; }
        public object CurrentView { get; set; }
        public bool CanCreateIssue { get; set; }
        public IRibbonUI RibbonUi { get; set; }

        public void Initialised(object context)
        {
            explorer = (Explorer)CurrentView;
            explorer.SelectionChange += ExplorerOnSelectionChange;
        }

        public void CreateIssue(IRibbonControl ribbonControl)
        {
            if (!CanCreateIssue) return;

            var issueViewModel = createIssuesViewModelFactory();
            issueViewModel.Initialise(selectedMailItem);
            contentHost.AddOrActivate(issueViewModel);
            CanCreateIssue = false;
        }

        public void ShowSettings(IRibbonControl ribbonControl)
        {
            contentHost.AddOrActivate(settingsViewModel);
        }

        private void ExplorerOnSelectionChange()
        {
            // If a single mail item is selected, we can create an issue for that mail item
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
                            CanCreateIssue = true;
                        }
                        else
                        {
                            CanCreateIssue = false;
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
                    CanCreateIssue = false;
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
            contentHost.RegisterSelf(register);
        }
    }
}
