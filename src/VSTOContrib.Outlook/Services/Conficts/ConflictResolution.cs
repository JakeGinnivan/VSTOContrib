namespace VSTOContrib.Outlook.Services.Conficts
{
    /// <summary>
    /// The resolution to the conflict. For remote win, set SaveLocal to the winning instance and it will be saved.
    /// For a merge set both as the merged instance, and both will be saved.
    /// </summary>
    /// <typeparam name="TType"></typeparam>
    public class ConflictResolution<TType>
    {
        /// <summary>
        /// Set SaveLocal to the item to save locally, if null nothing will be saved locally
        /// </summary>
        public TType SaveLocal { get; set; }
        /// <summary>
        /// Set SaveLocal to the item to save remotely, if null nothing will be saved remotely
        /// </summary>
        public TType SaveRemote { get; set; }
    }
}