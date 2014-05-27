using System;

namespace GitHubForOutlookAddin.Features
{
    public interface ITaskPaneContent
    {
        event Action OnClose;
    }
}