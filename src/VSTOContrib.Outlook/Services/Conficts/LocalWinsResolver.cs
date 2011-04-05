namespace VSTOContrib.Outlook.Services.Conficts
{
    /// <summary>
    /// Local wins conflict resolver.
    /// </summary>
    /// <typeparam name="TType">The type of the type.</typeparam>
    public class LocalWinsResolver<TType> : IConflictResolver<TType>
    {
        /// <summary>
        /// Resolves the specified local entry.
        /// </summary>
        /// <param name="localEntry">The local entry.</param>
        /// <param name="remoteEntry">The remote entry.</param>
        /// <returns></returns>
        public ConflictResolution<TType> Resolve(TType localEntry, TType remoteEntry)
        {
            return new ConflictResolution<TType> { SaveRemote = localEntry };
        }
    }
}