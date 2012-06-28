using System;
using Autofac;

namespace OutlookAutofacQuickStart
{
    public class AddinBootstrapper : IDisposable
    {
        private readonly IContainer container;

        public AddinBootstrapper()
        {
            var containerBuilder = new ContainerBuilder();

            RegisterComponents(containerBuilder);

            container = containerBuilder.Build();
        }

        private static void RegisterComponents(ContainerBuilder containerBuilder)
        {

        }

        public object Resolve(Type type)
        {
            return container.Resolve(type);
        }

        public T Resolve<T>()
        {
            return container.Resolve<T>();
        }

        public void Dispose()
        {
            container.Dispose();
        }
    }
}