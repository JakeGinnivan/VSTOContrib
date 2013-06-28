using System;
using Autofac;
using TwitterFeedOutlookAddin.Core.Services;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace TwitterFeedOutlookAddin.Core
{
    public class AddinBootstrapper : IDisposable
    {
        readonly IContainer container;

        public AddinBootstrapper()
        {
            var containerBuilder = new ContainerBuilder();

            containerBuilder.RegisterType<TwitterService>().As<ITwitterService>();
            containerBuilder.RegisterType<ContactFeed>()
                .As<IRibbonViewModel>()
                .AsSelf();
            container = containerBuilder.Build();
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