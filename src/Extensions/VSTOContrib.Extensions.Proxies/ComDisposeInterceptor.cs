using System.Runtime.InteropServices;
using Castle.DynamicProxy;

namespace VSTOContrib.Extensions.Proxies
{
    /// <summary>
    /// Dynamic Proxy interceptor which handles the Dispose method, and cleans up wrapped COM Object
    /// </summary>
    public class ComDisposeInterceptor : IInterceptor
    {
        /// <summary>
        /// Intercepts the specified invocation.
        /// </summary>
        /// <param name="invocation">The invocation.</param>
        public void Intercept(IInvocation invocation)
        {
            if (invocation.Method.Name == "Dispose")
            {
                var target = ((IProxyTargetAccessor)invocation.Proxy).DynProxyGetTarget();
                if (Marshal.IsComObject(target))
                {
                    Marshal.ReleaseComObject(target);
                }
            }
            else
            {
                invocation.Proceed();
            }
        }
    }
}
