using Microsoft.Office.Interop.Outlook;

namespace GitHubForOutlookAddin.Features.CreateIssue
{
    public interface ICreateIssueViewModel : ITaskPaneContent
    {
        void Initialise(MailItem selectedMailItem);
    }
}