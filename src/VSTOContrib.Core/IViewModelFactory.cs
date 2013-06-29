using System;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Core
{
    /// <summary>
    /// Creates instances of view models
    /// </summary>
    public interface IViewModelFactory
    {
        /// <summary>
        /// Builds the requested viewmodel type
        /// </summary>
        /// <param name="viewModelType"></param>
        /// <returns></returns>
        IRibbonViewModel Resolve(Type viewModelType);

        /// <summary>
        /// Releases the viewmodel instance and gives the factory the chance to clean up any related services
        /// </summary>
        /// <param name="viewModelInstance"></param>
        void Release(IRibbonViewModel viewModelInstance);
    }
}