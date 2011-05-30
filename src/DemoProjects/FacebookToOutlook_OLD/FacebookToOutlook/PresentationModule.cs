using Autofac;
using AutoMapper;
using FacebookToOutlook.Core;
using FacebookToOutlook.Presentation;
using FacebookToOutlook.Presentation.ViewModels;
using FacebookToOutlook.Presentation.ViewModels.ContactSync;
using FacebookToOutlook.ViewModels;

namespace FacebookToOutlook
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
