using IronGitHub.Entities;
using VSTOContrib.Core;

namespace GitHubForOutlookAddin.Features.CreateIssue
{
    public class RepositoryModel : NotifyPropertyChanged
    {
        readonly Repository repository;

        public RepositoryModel(Repository repository)
        {
            this.repository = repository;
        }

        public string Name { get { return repository.FullName; }}

        public Repository GitHubObject
        {
            get { return repository; }
        }
    }
}