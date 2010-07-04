using Autofac;

namespace FacebookToOutlook.Services
{
    public class ServicesModule : Module
    {
        protected override void Load(ContainerBuilder builder)
        {
            builder.Register(c => new DialogService()).As<IDialogService>();
            builder.RegisterType<OutlookMetaService>().As<IOutlookMetaService>();
        }
    }
}
