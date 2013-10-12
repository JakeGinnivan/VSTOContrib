using System;
using VSTOContrib.Core.RibbonFactory.Internal;

namespace GitHubForOutlook.Core.Features.Settings
{
    public interface ISettingsViewModel
    {
        void Init(ICustomTaskPaneWrapper settingsTaskPane);
        void LoginCallback(Action action);
    }
}