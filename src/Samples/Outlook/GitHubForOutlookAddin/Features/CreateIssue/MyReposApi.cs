using System.Collections.Generic;
using System.Net;
using System.Threading.Tasks;
using IronGitHub;
using IronGitHub.Entities;

namespace GitHubForOutlookAddin.Features.CreateIssue
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