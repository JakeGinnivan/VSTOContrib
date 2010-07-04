using System;

namespace Office.Utility
{
    /// <summary>
    /// Generic wrapper class to make any class disposable
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class AutoDispose<T> : IDisposable
    {
        private readonly T _resource;
        readonly Action<T> _managedCleanupAction;
        private bool _disposed;

        /// <summary>
        /// Initializes a new instance of the <see cref="AutoDispose&lt;T&gt;"/> class.
        /// </summary>
        /// <param name="resource">The resource.</param>
        /// <param name="managedCleanupAction">The resource cleanup action.</param>
        public AutoDispose(T resource, Action<T> managedCleanupAction)
        {
            _resource = resource;
            _managedCleanupAction = managedCleanupAction;
        }

        /// <summary>
        /// Gets the resource.
        /// </summary>
        /// <value>The resource.</value>
        public T Resource
        {
            get { return _resource; }
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Releases unmanaged and - optionally - managed resources
        /// </summary>
        /// <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _managedCleanupAction(Resource);
                }

                // There are no unmanaged resources to release, but
                // if we add them, they need to be released here.
            }
            _disposed = true;
        }
    }
}