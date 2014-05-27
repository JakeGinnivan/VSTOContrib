using Autofac;
using VSTOContrib.Autofac;
using WikipediaWordAddin.Services;
using WikipediaWordAddin.WpfControls;

namespace WikipediaWordAddin
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