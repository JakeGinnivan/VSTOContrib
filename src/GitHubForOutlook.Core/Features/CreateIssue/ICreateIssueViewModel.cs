using Microsoft.Office.Interop.Outlook;
using VSTOContrib.Core.RibbonFactory.Internal;

namespace GitHubForOutlook.Core.Features.CreateIssue
{
    public interface ICreateIssueViewModel
    {
        void CreateIssueFor(MailItem selectedMailItem);
        void Init(ICustomTaskPaneWrapper createIssueTaskPane);
    }
}