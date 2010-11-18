namespace Outlook.Utility.Services
{
    /// <summary>
    /// Interface resposenting a synchronisation service
    /// </summary>
    public interface ISynchronisationService
    {
        /// <summary>
        /// Gets a value indicating whether a sync is in progress.
        /// </summary>
        /// <value><c>true</c> if sync is in progress; otherwise, <c>false</c>.</value>
        /// 
        bool SyncInProgress { get; }

        /// <summary>
        /// Runs synchronisation
        /// </summary>
        void Synchronise();

        /// <summary>
        /// Runs synchronisation asynchronously
        /// </summary>
        void SynchroniseAsync();
    }
}
