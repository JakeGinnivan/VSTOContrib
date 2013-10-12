using Autofac;
using GitHubForOutlook.Core.Features.CreateIssue;
using GitHubForOutlook.Core.Features.Settings;
using IronGitHub;
using VSTOContrib.Autofac;

namespace GitHubForOutlook.Core
{
    public class AddinModule : Module
    {
        protected override void Load(ContainerBuilder builder)
        {
            builder.RegisterType<GitHubApi>().AsSelf().SingleInstance();
            builder.RegisterType<SettingsViewModel>().As<ISettingsViewModel>();
            builder.RegisterType<CreateIssueViewModel>().As<ICreateIssueViewModel>();

            builder.RegisterRibbonViewModels(typeof(AddinModule).Assembly);
        }
    }
}
