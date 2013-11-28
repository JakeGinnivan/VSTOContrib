namespace VSTOContrib.Outlook.Services
{
    /// <summary>
    /// Summary of Synchronisation Results
    /// </summary>
    public class SynchronisationResults
    {
        /// <summary>
        /// Gets or sets the message.
        /// </summary>
        /// <value>The message.</value>
        public string Message { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="SynchronisationResults"/> is success.
        /// </summary>
        /// <value><c>true</c> if success; otherwise, <c>false</c>.</value>
        public bool Success { get; set; }

        /// <summary>
        /// Gets or sets the number updated local.
        /// </summary>
        /// <value>The number updated local.</value>
        public int NumberUpdatedLocal { get; set; }

        /// <summary>
        /// Gets or sets the number updated remote.
        /// </summary>
        /// <value>The number updated remote.</value>
        public int NumberUpdatedRemote { get; set; }

        /// <summary>
        /// Gets or sets the number deleted remote.
        /// </summary>
        /// <value>The number deleted remote.</value>
        public int NumberDeletedRemote { get; set; }

        /// <summary>
        /// Gets or sets the number deleted local.
        /// </summary>
        /// <value>The number deleted local.</value>
        public int NumberDeletedLocal { get; set; }

        /// <summary>
        /// Gets or sets the number conflicts.
        /// </summary>
        /// <value>The number conflicts.</value>
        public int NumberConflicts { get; set; }
    }
}