using System.Reflection;
using Autofac;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Autofac
{
    /// <summary>
    /// Autofac Registration Extensions
    /// </summary>
    public static class RegistrationExtensions
    {
        /// <summary>
        /// Registers all Ribbon View Models with autofac in the given assembly
        /// </summary>
        /// <param name="containerBuilder"></param>
        /// <param name="assemblyToScan"></param>
        public static void RegisterRibbonViewModels(this ContainerBuilder containerBuilder, Assembly assemblyToScan)
        {
            containerBuilder.RegisterAssemblyTypes(assemblyToScan)
                .AssignableTo<IRibbonViewModel>()
                .AsSelf()
                .InstancePerLifetimeScope();
        }
    }
}