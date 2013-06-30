using Autofac;
using VSTOContrib.Autofac;
using WikipediaWordAddin.Core.Services;
using WikipediaWordAddin.Core.WpfControls;

namespace WikipediaWordAddin.Core
{
    public class AddinModule : Module
    {
        protected override void Load(ContainerBuilder builder)
        {
            builder.RegisterType<WikipediaService>().As<IWikipediaService>().InstancePerLifetimeScope();
            builder.RegisterRibbonViewModels(typeof(AddinModule).Assembly);
            builder.RegisterType<WikipediaResultsViewModel>().AsSelf().InstancePerLifetimeScope();
        }
    }
}