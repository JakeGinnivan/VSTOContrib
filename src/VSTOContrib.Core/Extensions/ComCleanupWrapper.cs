using System;
using System.Collections;
using System.Collections.Generic;

namespace VSTOContrib.Core.Extensions
{
    internal sealed class ComCleanupWrapper<T> : IEnumerable<T> where T : class
    {
        private readonly IEnumerable comCollection;
        private readonly Action<T> cleanup;
        private readonly Action<IEnumerator> enumeratorCleanup;
        private bool enumerated;

        public ComCleanupWrapper(IEnumerable comCollection, Action<T> cleanup, Action<IEnumerator> enumeratorCleanup)
        {
            this.comCollection = comCollection;
            this.cleanup = cleanup;
            this.enumeratorCleanup = enumeratorCleanup;
        }

        public IEnumerator<T> GetEnumerator()
        {
            //This enumerator cleans up items as it is enumerated, so we need to stop multiple enumerations.
            if (enumerated)
                throw new InvalidOperationException("Can only enumerate collection once");
            enumerated = true;

            return new ComCleanupEnumerator<T>(comCollection.GetEnumerator(), cleanup, enumeratorCleanup);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}