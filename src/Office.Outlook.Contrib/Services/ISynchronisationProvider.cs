using System;
using System.Collections.Generic;

namespace Office.Outlook.Contrib.Services
{
    /// <summary>
    /// Synchronisation Provider for the Generic Synchronisation service
    /// </summary>
    /// <typeparam name="TType">The type of the type.</typeparam>
    /// <typeparam name="TKey">The type of the key.</typeparam>
    public interface ISynchronisationProvider<TType, TKey>
    {
        /// <summary>
        /// Gets the modified entries.
        /// </summary>
        /// <param name="lastSync">The last sync.</param>
        /// <returns></returns>
        IEnumerable<TType> GetModifiedEntries(DateTime? lastSync);

        /// <summary>
        /// Gets the deleted entries.
        /// You can use the DeletedEventsHelper to calculate deleted events vs 
        /// items that have fallen out of scope (past events) which are not deleted.
        /// </summary>
        /// <param name="lastSync">The last sync.</param>
        /// <returns></returns>
        IEnumerable<TKey> GetDeletedEntries(DateTime? lastSync);

        /// <summary>
        /// Saves the entries.
        /// </summary>
        /// <param name="entries">The entries.</param>
        void SaveEntries(IEnumerable<TType> entries);

        /// <summary>
        /// Deletes the entries.
        /// </summary>
        /// <param name="keys">The keys.</param>
        void DeleteEntries(IEnumerable<TKey> keys);
    }
}
