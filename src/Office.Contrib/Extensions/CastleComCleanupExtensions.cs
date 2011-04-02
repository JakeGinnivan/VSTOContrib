using System;
using Castle.DynamicProxy;

namespace Office.Contrib.Extensions
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
        public static T1 WithComCleanup<T, T1>(this T resource)
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

    /// <summary>
    /// Interceptor for the .Resource get property
    /// </summary>
    public class ResourceInterceptor : IInterceptor
    {
        /// <summary>
        /// Intercepts the specified invocation.
        /// </summary>
        /// <param name="invocation">The invocation.</param>
        public void Intercept(IInvocation invocation)
        {
            if (invocation.Method.Name == "get_Resource")
            {
                invocation.ReturnValue = ((IProxyTargetAccessor)invocation.Proxy).DynProxyGetTarget();
            }
            else
            {
                invocation.Proceed();
            }
        }
    }
}
