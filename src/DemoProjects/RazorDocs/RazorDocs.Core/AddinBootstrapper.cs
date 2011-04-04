using System;
using Autofac;

namespace RazorDocs.Core
{
    //Todo create a AutofacAddinBase
    public class AddinBootstrapper : IDisposable
    {
        private readonly IContainer _container;

        public AddinBootstrapper()
        {
            var containerBuilder = new ContainerBuilder();

            containerBuilder.RegisterType<RazorDocs>();

            _container = containerBuilder.Build();
        }

        public object Resolve(Type type)
        {
            return _container.Resolve(type);
        }

        public T Resolve<T>()
        {
            return _container.Resolve<T>();
        }

        public void Dispose()
        {
            _container.Dispose();
        }
    }
}
 