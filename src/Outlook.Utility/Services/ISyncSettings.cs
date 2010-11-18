using System;

namespace Outlook.Utility.Services
{
    /// <summary>
    /// Synchronisation settings
    /// </summary>
    public interface ISyncSettings
    {
        /// <summary>
        /// Gets or sets the time of the last sucessful synchronisation
        /// </summary>
        /// <value>The last sync.</value>
        DateTime? LastSync { get; set; }

        /// <summary>
        /// Saves the settings
        /// </summary>
        void Save();
    }
}