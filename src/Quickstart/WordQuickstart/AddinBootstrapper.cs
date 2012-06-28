using System;

namespace WordQuickstart
{
    /// <summary>
    /// Move me to a new class library
    /// </summary>
    public class AddinBootstrapper : IDisposable
    {
        public object Resolve(Type type)
        {
            return Activator.CreateInstance(type);
        }

        public T Resolve<T>()
        {
            return (T)Resolve(typeof(T));
        }

        public void Dispose()
        {
        }
    }
}