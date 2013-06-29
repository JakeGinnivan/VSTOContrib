using Autofac;
using VSTOContrib.Autofac;
using WikipediaWordAddin.Core.Services;

namespace WikipediaWordAddin.Core
{
    public class AddinModule : Module
    {
        protected override void Load(ContainerBuilder builder)
        {
            builder.RegisterType<WikipediaService>().As<IWikipediaService>().InstancePerLifetimeScope();
            builder.RegisterRibbonViewModels(typeof(AddinModule).Assembly);
        }
    }
}