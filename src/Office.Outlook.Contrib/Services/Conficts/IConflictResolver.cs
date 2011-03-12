namespace Office.Outlook.Contrib.Services.Conficts
{
    /// <summary>
    /// Synchronisation Conflict resolver
    /// </summary>
    /// <typeparam name="TType">The type of the type.</typeparam>
    public interface IConflictResolver<TType>
    {
        /// <summary>
        /// Resolves the conflict
        /// </summary>
        /// <param name="localEntry">The local entry.</param>
        /// <param name="remoteEntry">The remote entry.</param>
        /// <returns>Conflict Resolution Outcome</returns>
        ConflictResolution<TType> Resolve(TType localEntry, TType remoteEntry);
    }
}