using System.Collections.Generic;
using System.Net;
using System.Threading.Tasks;
using IronGitHub;
using IronGitHub.Entities;

namespace GitHubForOutlook.Core.Features.CreateIssue
{
    public class MyReposApi : GitHubApiBase
    {
        public MyReposApi(GitHubApiContext context) : base(context)
        {
             
        }

        public async Task<IEnumerable<Repository>> GetMyRepositories()
        {
            HttpWebRequest request = CreateRequest("/user/repos?sort=updated");
            var apiResponse = await Complete<IEnumerable<Repository>>(request);
            return apiResponse.Result;
        }
    }
}