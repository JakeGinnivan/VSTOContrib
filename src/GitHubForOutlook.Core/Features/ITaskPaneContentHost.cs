using GitHubForOutlook.Core.Features.CreateIssue;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace GitHubForOutlook.Core.Features
{
    public interface ITaskPaneContentHost
    {
        void RegisterSelf(Register register);
        void AddOrActivate(ITaskPaneContent taskPaneContent);
    }
}