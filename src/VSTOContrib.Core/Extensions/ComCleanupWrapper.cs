using System;
using System.Collections;
using System.Collections.Generic;

namespace VSTOContrib.Core.Extensions
{
    internal sealed class ComCleanupWrapper<T> : IEnumerable<T> where T : class
    {
        private readonly IEnumerable _comCollection;
        private readonly Action<T> _cleanup;
        private readonly Action<IEnumerator> _enumeratorCleanup;
        private bool _enumerated;

        public ComCleanupWrapper(IEnumerable comCollection, Action<T> cleanup, Action<IEnumerator> enumeratorCleanup)
        {
            _comCollection = comCollection;
            _cleanup = cleanup;
            _enumeratorCleanup = enumeratorCleanup;
        }

        public IEnumerator<T> GetEnumerator()
        {
            //This enumerator cleans up items as it is enumerated, so we need to stop multiple enumerations.
            if (_enumerated)
                throw new InvalidOperationException("Can only enumerate collection once");
            _enumerated = true;

            return new ComCleanupEnumerator<T>(_comCollection.GetEnumerator(), _cleanup, _enumeratorCleanup);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}