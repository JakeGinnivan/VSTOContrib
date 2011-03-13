using System;
using System.Collections;
using System.Collections.Generic;

namespace Office.Contrib
{
    /// <summary>
    /// Custom Enumerator which will release the previous COM object when MoveNext is called
    /// </summary>
    /// <typeparam name="T"></typeparam>
    internal sealed class ComCleanupEnumerator<T> : IEnumerator<T> where T : class
    {
        private IEnumerator _source;
        private readonly Action<T> _cleanup;
        private readonly Action<IEnumerator> _cleanupEnumerator;

        /// <summary>
        /// Initializes a new instance of the <see cref="ComCleanupEnumerator&lt;T&gt;"/> class.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="cleanup">The cleanup.</param>
        public ComCleanupEnumerator(IEnumerator source, Action<T> cleanup) :
            this(source, cleanup, c=> { })
        { }

        /// <summary>
        /// Initializes a new instance of the <see cref="ComCleanupEnumerator&lt;T&gt;"/> class.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="cleanup">The cleanup ection.</param>
        /// <param name="cleanupEnumerator">The cleanup enumerator action.</param>
        public ComCleanupEnumerator(IEnumerator source, Action<T> cleanup, Action<IEnumerator> cleanupEnumerator)
        {
            _source = source;
            _cleanup = cleanup;
            _cleanupEnumerator = cleanupEnumerator;
        }

        /// <summary>
        /// Gets the current item.
        /// </summary>
        /// <value>The current item.</value>
        public T Current { get; private set; }

        /// <summary>
        /// Gets the current item.
        /// </summary>
        /// <value>The current item.</value>
        object IEnumerator.Current
        {
            get { return Current; }
        }

        /// <summary>
        /// Advances the enumerator to the next element of the collection.
        /// </summary>
        /// <returns>
        /// true if the enumerator was successfully advanced to the next element; false if the enumerator has passed the end of the collection.
        /// </returns>
        /// <exception cref="T:System.InvalidOperationException">
        /// The collection was modified after the enumerator was created.
        /// </exception>
        public bool MoveNext()
        {
            if (null != _source)
            {
                _cleanup(Current);
                Current = null;
                //Cannot call Current if enumeration is finished
                if (!_source.MoveNext())
                    return false;
                var current = _source.Current as T;
                if (current == null)
                    return MoveNext();
                Current = current;
            }

            return true;
        }

        /// <summary>
        /// Sets the enumerator to its initial position, which is before the first element in the collection.
        /// </summary>
        /// <exception cref="T:System.InvalidOperationException">
        /// The collection was modified after the enumerator was created.
        /// </exception>
        public void Reset()
        {
            if (null == _source) throw new ObjectDisposedException("ComCleanupEnumerator");

            _cleanup(Current);
            _source.Reset();
            Current = null;
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            var source = _source as IDisposable;
            if (source != null)
                source.Dispose();

            //Cleanup current if there is still a value (Dispose called before MoveNext returns false)
            if (Current != null)
                _cleanup(Current);

            if (_cleanupEnumerator !=null)
                _cleanupEnumerator(_source);
            _source = null;
        }
    }
}