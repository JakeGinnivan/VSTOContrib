using Castle.DynamicProxy;

namespace VSTOContrib.Extensions.Proxies
{
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