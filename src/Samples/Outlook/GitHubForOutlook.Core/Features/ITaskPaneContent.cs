using System;

namespace GitHubForOutlook.Core.Features
{
    public interface ITaskPaneContent
    {
        event Action OnClose;
    }
}