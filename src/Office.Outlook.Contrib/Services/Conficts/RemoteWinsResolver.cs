namespace Office.Outlook.Contrib.Services.Conficts
{
    /// <summary>
    /// Remote wins conflict resolver.. Server always wins
    /// </summary>
    /// <typeparam name="TType">The type of the type.</typeparam>
    public class RemoteWinsResolver<TType> : IConflictResolver<TType>
    {
        /// <summary>
        /// Resolves the specified conflict.
        /// </summary>
        /// <param name="localEntry">The local entry.</param>
        /// <param name="remoteEntry">The remote entry.</param>
        /// <returns></returns>
        public ConflictResolution<TType> Resolve(TType localEntry, TType remoteEntry)
        {
            return new ConflictResolution<TType> { SaveLocal = remoteEntry };
        }
    }
}