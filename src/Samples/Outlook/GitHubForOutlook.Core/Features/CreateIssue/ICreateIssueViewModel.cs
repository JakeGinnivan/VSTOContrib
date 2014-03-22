using Microsoft.Office.Interop.Outlook;

namespace GitHubForOutlook.Core.Features.CreateIssue
{
    public interface ICreateIssueViewModel : ITaskPaneContent
    {
        void Initialise(MailItem selectedMailItem);
    }
}