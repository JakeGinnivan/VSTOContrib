using System;

namespace VSTOContrib.Core.Extensions
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ComObjectWrapper<T> : IDisposable where T : class
    {
        /// <summary>
        /// Wrapped Resource
        /// </summary>
        public T Resource { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ComObjectWrapper&lt;T&gt;"/> class.
        /// </summary>
        /// <param name="resource">The resource.</param>
        public ComObjectWrapper(T resource)
        {
            Resource = resource;
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Resource.ReleaseComObject();
        }
    }
}