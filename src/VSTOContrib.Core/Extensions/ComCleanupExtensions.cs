using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace VSTOContrib.Core.Extensions
{
    /// <summary>
    /// Extension methods which help a deterministic cleanup model
    /// </summary>
    public static class ComCleanupExtensions
    {
        /// <summary>
        /// Wraps the Com resource in an IDisposable proxy which releases 
        /// the Com object when Dispose is called.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="resource">The resource.</param>
        /// <returns></returns>
        public static ComObjectWrapper<T> WithComCleanup<T>(this T resource)
            where T : class
        {
            return new ComObjectWrapper<T>(resource);
        }

        /// <summary>
        /// Enables Linq on any COM collection. Releases each iterated item deterministically
        /// as the collection is enumerated
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="comCollection">The COM collection.</param>
        /// <returns></returns>
        public static IEnumerable<T> ComLinq<T>(this IEnumerable comCollection)
            where T : class
        {
            return new ComCleanupWrapper<T>(comCollection, ReleaseComObject, ReleaseComObject);
        }

        /// <summary>
        /// Releases the COM object.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="resource">The Com object to releases.</param>
        public static void ReleaseComObject<T>(this T resource) where T : class
        {
            if (resource != null && Marshal.IsComObject(resource))
                Marshal.ReleaseComObject(resource);
        }
    }
}
