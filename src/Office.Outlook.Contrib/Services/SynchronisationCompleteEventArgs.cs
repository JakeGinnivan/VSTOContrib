using System;

namespace Office.Outlook.Contrib.Services
{
    /// <summary>
    /// Event arguements when synchronisation is complete
    /// </summary>
    public class SynchronisationCompleteEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SynchronisationCompleteEventArgs"/> class.
        /// </summary>
        /// <param name="results">The results.</param>
        public SynchronisationCompleteEventArgs(SynchronisationResults results)
        {
            SynchronisationResults = results;
        }

        /// <summary>
        /// Gets or sets the synchronisation results.
        /// </summary>
        /// <value>The synchronisation results.</value>
        public SynchronisationResults SynchronisationResults { get; set; }
    }
}