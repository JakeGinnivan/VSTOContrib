using Autofac;
using AutoMapper;
using FacebookToOutlook.Core;
using FacebookToOutlook.Presentation.ViewModels;
using FacebookToOutlook.Presentation.ViewModels.ContactSync;

namespace FacebookToOutlook.Presentation
{
    public class PresentationModule : Module
    {
        protected override void Load(ContainerBuilder builder)
        {
            builder.RegisterType<ConfigurationViewModel>();
            builder.RegisterType<EventConfigurationViewModel>();
            builder.RegisterType<ContactConfigurationViewModel>();
            builder.RegisterType<ContactSyncSetupViewModel>();
            builder.RegisterType<ContactListsBuilder>();
            builder.RegisterType<UnmatchedContactsViewModel>();
            builder.RegisterType<ContactSync>();

            SetupMappings();
        }

        private static void SetupMappings()
        {
            Mapper.CreateMap<IEventConfigurationSettings, EventConfigurationViewModel>();
            Mapper.CreateMap<EventConfigurationViewModel, IEventConfigurationSettings>();
            Mapper.CreateMap<IContactConfigurationSettings, ContactConfigurationViewModel>();
            Mapper.CreateMap<ContactConfigurationViewModel, IContactConfigurationSettings>();
        }
    }
}
