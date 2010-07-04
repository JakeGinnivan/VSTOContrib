using System;
using System.Collections;
using System.Collections.Generic;

namespace Office.Utility
{
    internal sealed class ComCleanupEnumerator<T> : IEnumerator<T> where T : class
    {
        private IEnumerator _source;
        private readonly Action<T> _cleanup;
        private readonly Action<IEnumerator> _cleanupEnumerator;

        public ComCleanupEnumerator(IEnumerator source, Action<T> cleanup) :
            this(source, cleanup, c=> { })
        { }

        public ComCleanupEnumerator(IEnumerator source, Action<T> cleanup, Action<IEnumerator> cleanupEnumerator)
        {
            _source = source;
            _cleanup = cleanup;
            _cleanupEnumerator = cleanupEnumerator;
        }

        public T Current { get; private set; }

        object IEnumerator.Current
        {
            get { return Current; }
        }

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

        public void Reset()
        {
            if (null == _source) throw new ObjectDisposedException("ComCleanupEnumerator");

            _cleanup(Current);
            _source.Reset();
            Current = null;
        }

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