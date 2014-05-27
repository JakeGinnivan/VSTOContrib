using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace GitHubForOutlookAddin.Features
{
    public interface ITaskPaneContentHost
    {
        void RegisterSelf(Register register);
        void AddOrActivate(ITaskPaneContent taskPaneContent);
    }
}