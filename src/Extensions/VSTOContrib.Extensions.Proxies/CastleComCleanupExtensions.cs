using System;

namespace VSTOContrib.Extensions.Proxies
{
    /// <summary>
    /// Contains extensions which rely on castle proxies
    /// </summary>
    public static class CastleComCleanupExtensions
    {
        private static readonly ComProxyGenerator ComProxyGenerator;

        static CastleComCleanupExtensions()
        {
            ComProxyGenerator = new ComProxyGenerator();            
        }

        /// <summary>
        /// Wraps the Com resource in an IDisposable proxy which releases 
        /// the Com object when Dispose is called.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="T1">The type of the 1.</typeparam>
        /// <param name="resource">The resource.</param>
        /// <returns></returns>
        public static T1 WithComCleanupProxy<T, T1>(this T resource)
            where T1 : T, IDisposable
            where T : class
        {
            if (resource == null) return (T1)(object)null;
            return ComProxyGenerator
                .CreateComProxy<T, T1>(
                    resource,
                    new ComDisposeInterceptor(),
                    new ResourceInterceptor());
        }
    }
}
