using System;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Core
{
    /// <summary>
    /// Creates instances of view models using Activator.CreateInstance (viewmodels need a default constructor
    /// </summary>
    public class DefaultViewModelFactory : IViewModelFactory
    {
        /// <summary>
        /// Builds the requested viewmodel type
        /// </summary>
        /// <param name="viewModelType"></param>
        /// <returns></returns>
        public IRibbonViewModel Resolve(Type viewModelType)
        {
            return (IRibbonViewModel) Activator.CreateInstance(viewModelType);
        }

        /// <summary>
        /// Releases the viewmodel instance and gives the factory the chance to clean up any related services
        /// </summary>
        /// <param name="viewModelInstance"></param>
        public void Release(IRibbonViewModel viewModelInstance)
        {
            // ReSharper disable once SuspiciousTypeConversion.Global
            var disposable = viewModelInstance as IDisposable;
            if (disposable != null)
                disposable.Dispose();
        }
    }
}