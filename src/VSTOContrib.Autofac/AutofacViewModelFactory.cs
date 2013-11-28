using System;
using System.Collections.Generic;
using Autofac;
using Autofac.Core;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Autofac
{
    /// <summary>
    /// Autofac View Model Factory
    /// </summary>
    public class AutofacViewModelFactory : IViewModelFactory, IDisposable
    {
        readonly IContainer autofacContainer;

        readonly Dictionary<IRibbonViewModel, ILifetimeScope> lifetimeScopeLookups = 
            new Dictionary<IRibbonViewModel, ILifetimeScope>();

        readonly bool shouldDisposeContainer;

        /// <summary>
        /// Creates new instance of AutofacViewModelFactory using an Autofac Module
        /// 
        /// This will cause the container to be disposed when the AutofacViewModelFactory is disposed
        /// </summary>
        public AutofacViewModelFactory(IModule moduleToRegister)
        {
            var containerBuilder = new ContainerBuilder();
            containerBuilder.RegisterModule(moduleToRegister);
            autofacContainer = containerBuilder.Build();
            shouldDisposeContainer = true;
        }

        /// <summary>
        /// Creates new instance of AutofacViewModelFactory using an Autofac Container
        /// 
        /// This will cause the container NOT to be disposed when the AutofacViewModelFactory is disposed
        /// </summary>
        /// <param name="autofacContainer"></param>
        public AutofacViewModelFactory(IContainer autofacContainer)
        {
            this.autofacContainer = autofacContainer;
        }

        /// <summary>
        /// Builds the requested viewmodel type
        /// </summary>
        /// <returns></returns>
        public IRibbonViewModel Resolve(Type viewModelType)
        {
            var lifetime = autofacContainer.BeginLifetimeScope();
            var viewModel = (IRibbonViewModel)lifetime.Resolve(viewModelType);
            lifetimeScopeLookups.Add(viewModel, lifetime);

            return viewModel;
        }

        /// <summary>
        /// Releases the viewmodel instance and gives the factory the chance to clean up any related services
        /// </summary>
        /// <param name="viewModelInstance"></param>
        public void Release(IRibbonViewModel viewModelInstance)
        {
            var lifetimeScope = lifetimeScopeLookups[viewModelInstance];
            lifetimeScopeLookups.Remove(viewModelInstance);
            lifetimeScope.Dispose();
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (shouldDisposeContainer)
                autofacContainer.Dispose();
        }
    }
}
