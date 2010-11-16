using System;
using System.Collections.Generic;
using System.Linq;
using Castle.DynamicProxy;

namespace Office.Utility
{
    /// <summary>
    /// 
    /// </summary>
    public class ComProxyGenerator : ProxyGenerator
    {
        /// <summary>
        /// Creates the COM proxy.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="T1">The type of the 1.</typeparam>
        /// <param name="comObject">The COM object.</param>
        /// <param name="interceptors">The interceptors.</param>
        /// <returns></returns>
        public T1 CreateComProxy<T, T1>(T comObject, params IInterceptor[] interceptors) where T1 : T, IDisposable
        {

            var options = ProxyGenerationOptions.Default;
            var type = CreateInterfaceProxyTypeWithTargetInterface(typeof(T), new[] { typeof(T1) }, options);

            var ctorArgs = GetConstructorArguments(comObject, interceptors, options);
            return (T1)Activator.CreateInstance(type, ctorArgs.ToArray());
        }

        private static IEnumerable<object> GetConstructorArguments(object target, IInterceptor[] interceptors,
                                                        ProxyGenerationOptions options)
        {
            // create constructor arguments (initialized with mixin implementations, 
            // interceptors and target type constructor arguments)
            var arguments = new List<object>(options.MixinData.Mixins) { interceptors, target };
            if (options.Selector != null)
            {
                arguments.Add(options.Selector);
            }
            return arguments;
        }
    }
}
