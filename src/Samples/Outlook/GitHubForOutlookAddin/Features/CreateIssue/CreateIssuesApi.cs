using System.Threading.Tasks;
using IronGitHub;
using IronGitHub.Entities;

namespace GitHubForOutlookAddin.Features.CreateIssue
{
    public class CreateIssuesApi : GitHubApiBase
    {
        public CreateIssuesApi(GitHubApiContext context) : base(context)
        {
        }

        public async Task<Issue> CreateIssue(string repoFullName, string issueTitle, string issueDescription)
        {
            var request = CreateRequest(string.Format("/repos/{0}/issues", repoFullName));

            var response = await PostAsJson<NewIssue, Issue>(request, new NewIssue
            {
                Title = issueTitle,
                Body = issueDescription
            });

            return response.Result;

        }
    }
}