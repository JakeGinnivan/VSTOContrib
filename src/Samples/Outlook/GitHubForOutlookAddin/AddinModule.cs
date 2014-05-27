using Autofac;
using GitHubForOutlookAddin.Features;
using GitHubForOutlookAddin.Features.CreateIssue;
using GitHubForOutlookAddin.Features.Settings;
using IronGitHub;
using VSTOContrib.Autofac;

namespace GitHubForOutlookAddin
{
    public class AddinModule : Module
    {
        protected override void Load(ContainerBuilder builder)
        {
            builder.RegisterType<GitHubApi>().AsSelf().SingleInstance();
            builder.RegisterType<SettingsViewModel>().As<ISettingsViewModel>();
            builder.RegisterType<CreateIssueViewModel>().As<ICreateIssueViewModel>();
            builder.RegisterType<GitHubSettings>().As<IGitHubSettings>().SingleInstance();
            builder.RegisterType<TaskPaneContentHost>().As<ITaskPaneContentHost>();

            builder.RegisterRibbonViewModels(typeof(AddinModule).Assembly);
        }
    }
}
