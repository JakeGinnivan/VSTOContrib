using Autofac;
using FacebookToOutlook.Core;
using FacebookToOutlook.Data;
using FacebookToOutlook.Properties;
using FacebookToOutlook.Services;
using Office.Outlook.Contrib.Services;

namespace FacebookToOutlook
{
    class SettingsModule : Module
    {
        protected override void Load(ContainerBuilder builder)
        {
            builder.Register(c => Settings.Default).As<ISyncSettings>();
            builder.Register(c => Settings.Default).As<IConfigurationSettings>();
            builder.Register(c => Settings.Default).As<IEventConfigurationSettings>();
            builder.Register(c => Settings.Default).As<IContactConfigurationSettings>();
            builder.Register(c => Settings.Default).As<IApplicationSettings>();
            builder.Register(c => Settings.Default).As<ISynchronisedEventInfo>();
        }
    }
}
