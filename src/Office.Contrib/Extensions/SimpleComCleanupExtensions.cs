using System;

namespace Office.Contrib.Extensions
{
    /// <summary>
    /// Contains extensions which rely on castle proxies
    /// </summary>
    public static class SimpleComCleanupExtensions
    {
        /// <summary>
        /// Wraps the Com resource in an IDisposable proxy which releases 
        /// the Com object when Dispose is called.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="resource">The resource.</param>
        /// <returns></returns>
        public static ComObjectWrapper<T> WithComCleanup<T>(this T resource)
            where T : class
        {
            return new ComObjectWrapper<T>(resource);
        }
    }

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
