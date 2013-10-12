﻿using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using IronGitHub;
using IronGitHub.Entities;
using Microsoft.Office.Interop.Outlook;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Outlook;

namespace GitHubForOutlook.Core.Features.CreateIssue
{
    public class CreateIssueViewModel : NotifyPropertyChanged, ICreateIssueViewModel
    {
        ICustomTaskPaneWrapper taskPane;
        readonly GitHubApi githubApi;
        MailItem currentMailItem;

        public CreateIssueViewModel(GitHubApi githubApi)
        {
            this.githubApi = githubApi;
            Repositories = new ObservableCollection<RepositoryModel>();
            CreateIssueCommand = new DelegateCommand(CreateIssue);
        }

        async void CreateIssue()
        {
            CreatingIssue = true;
            try
            {
                var createdIssue = await new CreateIssuesApi(githubApi.Context).CreateIssue(SelectedRepository, IssueTitle, IssueDescription);
                currentMailItem.UserProperties.SetPropertyValue("GitHubIssueId", OlUserPropertyType.olNumber, createdIssue.Id, false);
                taskPane.Visible = false;
            }
            finally
            {
                CreatingIssue = false;                
            }
        }

        public ObservableCollection<RepositoryModel> Repositories { get; private set; }

        public string SelectedRepository { get; set; }
        public string IssueTitle { get; set; }
        public string IssueDescription { get; set; }
        public bool CreatingIssue { get; private set; }

        public ICommand CreateIssueCommand { get; private set; }

        public async void CreateIssueFor(MailItem selectedMailItem)
        {
            IssueTitle = selectedMailItem.Subject;
            IssueDescription = string.Format("Reported by email from {0} ({1}) at {2}\r\n\r\n{3}",
                selectedMailItem.SenderName,
                selectedMailItem.SenderEmailAddress.Replace("@", " at ").Replace(".", " dot "),
                selectedMailItem.ReceivedTime.ToString("U"),
                selectedMailItem.Body);
            currentMailItem = selectedMailItem;

            if (Repositories.Count > 0) return;

            if (githubApi.Context.Authorization == null || githubApi.Context.Authorization == Authorization.Anonymous)
            {
                githubApi.Context.Authorize(new Authorization
                {
                    Id = Properties.Settings.Default.AuthorisationId,
                    Token = Properties.Settings.Default.AuthToken
                });
            }
            var repos = await new MyReposApi(githubApi.Context).GetMyRepositories();
            await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
            {
                foreach (var repository in repos.OrderBy(r => r.FullName))
                {
                    Repositories.Add(new RepositoryModel(repository));
                }
            });
        }

        public void Init(ICustomTaskPaneWrapper createIssueTaskPane)
        {
            taskPane = createIssueTaskPane;
        }
    }
}