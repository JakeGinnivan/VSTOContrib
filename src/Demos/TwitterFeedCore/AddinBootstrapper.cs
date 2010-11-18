using System;
using Autofac;
using Outlook.Utility.RibbonFactory;
using TwitterFeedCore.Services;
using TwitterFeedCore.TwitterFeed;

namespace TwitterFeedCore
{
    public class AddinBootstrapper : IDisposable
    {
        private readonly IContainer _container;

        public AddinBootstrapper()
        {
            var containerBuilder = new ContainerBuilder();

            containerBuilder.RegisterType<TwitterService>().As<ITwitterService>();
            containerBuilder.RegisterType<ContactFeed>()
                .As<IRibbonViewModel>()
                .AsSelf();
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
