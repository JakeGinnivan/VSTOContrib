using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using GitHubForOutlook.Core.Features.Settings;
using IronGitHub;
using IronGitHub.Entities;
using Microsoft.Office.Interop.Outlook;
using VSTOContrib.Core;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Outlook;
using Action = System.Action;

namespace GitHubForOutlook.Core.Features.CreateIssue
{
    public class CreateIssueViewModel : NotifyPropertyChanged, ICreateIssueViewModel
    {
        readonly GitHubApi githubApi;
        readonly IGitHubSettings settings;
        MailItem currentMailItem;

        public CreateIssueViewModel(GitHubApi githubApi, IGitHubSettings settings)
        {
            this.githubApi = githubApi;
            this.settings = settings;
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
                var reply = currentMailItem.ReplyAll();
                reply.Body = string.Format("I have created an issue at {0} to track this issue.\r\n\r\nThanks for reporting it.", 
                    createdIssue.HtmlUrl);
                reply.Display(Modal: false);

                OnClose();
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

        public async void Initialise(MailItem selectedMailItem)
        {
            IssueTitle = selectedMailItem.Subject;
            IssueDescription = string.Format("Reported by email from {0} at {1}\r\n\r\n{2}",
                selectedMailItem.SenderName,
                selectedMailItem.ReceivedTime.ToString("U"),
                selectedMailItem.Body);
            currentMailItem = selectedMailItem;

            if (Repositories.Count > 0) return;

            if (githubApi.Context.Authorization == null || githubApi.Context.Authorization == Authorization.Anonymous)
            {
                githubApi.Context.Authorize(new Authorization
                {
                    Id = settings.AuthorisationId,
                    Token = settings.AuthToken
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

        public event Action OnClose = () => { };
    }
}