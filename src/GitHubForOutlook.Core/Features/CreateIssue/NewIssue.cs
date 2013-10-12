using System.Runtime.Serialization;

namespace GitHubForOutlook.Core.Features.CreateIssue
{
    [DataContract]
    public class NewIssue
    {
        [DataMember(Name = "title")]
        public string Title { get; set; }

        [DataMember(Name = "body")]
        public string Body { get; set; }
    }
}