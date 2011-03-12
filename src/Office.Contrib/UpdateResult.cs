namespace Office.Contrib
{
    /// <summary>
    /// Deployment update result
    /// </summary>
    public class UpdateResult
    {
        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="UpdateResult"/> is success.
        /// </summary>
        /// <value><c>true</c> if success; otherwise, <c>false</c>.</value>
        public bool Success { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="UpdateResult"/> is updated.
        /// </summary>
        /// <value><c>true</c> if updated; otherwise, <c>false</c>.</value>
        public bool Updated { get; set; }
        /// <summary>
        /// Gets or sets the message.
        /// </summary>
        /// <value>The message.</value>
        public string Message { get; set; }
    }
}