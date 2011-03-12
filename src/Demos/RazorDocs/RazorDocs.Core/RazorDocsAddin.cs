using System;
using Autofac;

namespace RazorDocs.Core
{
    //Todo create a AutofacAddinBase
    public class RazorDocsAddin : IDisposable
    {
        private readonly IContainer _container;

        public RazorDocsAddin()
        {
            var containerBuilder = new ContainerBuilder();

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
 