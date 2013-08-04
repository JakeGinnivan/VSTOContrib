using Autofac;
using QuoteGeneratorAddin.Core.OfficeContexts;
using VSTOContrib.Autofac;

namespace QuoteGeneratorAddin.Core
{
    public class AddinModule : Module
    {
        protected override void Load(ContainerBuilder builder)
        {
            builder.RegisterRibbonViewModels(typeof(AddinModule).Assembly);
            builder.RegisterType<QuotesService>().As<IQuotesService>().SingleInstance();
        }
    }
}