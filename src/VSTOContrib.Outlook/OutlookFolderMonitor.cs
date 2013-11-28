using System;
using Microsoft.Office.Interop.Outlook;
using VSTOContrib.Core.Extensions;

namespace VSTOContrib.Outlook
{
    /// <summary>
    /// Monitors a MAPI Folder for item changes
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class OutlookFolderMonitor<T> : IDisposable
    {
        private MAPIFolder _deletedItemsFolder;
        private MAPIFolder _folder;
        private Items _items;
        private bool _disposed;

        /// <summary>
        /// Occurs when item is added.
        /// </summary>
        public event EventHandler<EventArgs<T>> ItemAdded;
        /// <summary>
        /// Occurs when item is modified or added.
        /// </summary>
        public event EventHandler<EventArgs<T>> ItemModified;
        /// <summary>
        /// Occurs when item is being deleted. Allows cancellation.
        /// </summary>
        public event EventHandler<CancelEventArgs<T>> ItemDeleting;

        /// <summary>
        /// Initializes a new instance of the <see cref="OutlookFolderMonitor&lt;T&gt;"/> class.
        /// </summary>
        /// <param name="session">The session.</param>
        /// <param name="folder">The folder to monitor.</param>
        public OutlookFolderMonitor(_NameSpace session, MAPIFolder folder)
        {
            _deletedItemsFolder = session.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
            HookupFolderEvents(folder);
        }

        private void HookupFolderEvents(MAPIFolder folder)
        {
            // Store a reference to the folder and to the items collection 
            // so that it remains alive for as long as we want. This keeps 
            // the ref count up on the underlying COM object and prevents
            // it from being intermittently released (then the events don't 
            // get fired).
            _items = folder.Items;
            _folder = folder;

            // Add listeners for the events we need.
            ((MAPIFolderEvents_12_Event)folder).BeforeItemMove += BeforeItemMove;
            _items.ItemChange += ItemsItemChange;
            _items.ItemAdd += ItemsItemAdd;
        }

        private void ItemsItemAdd(object item)
        {
            if (!(item is T) || ItemAdded == null) return;

            ItemAdded(this, new EventArgs<T>((T)item));
        }

        private void ItemsItemChange(object item)
        {
            if (!(item is T) || ItemModified == null) return;

            ItemModified(this, new EventArgs<T>((T)item));
        }

        private void BeforeItemMove(object item, MAPIFolder moveToFolder, ref bool cancel)
        {
            if (((moveToFolder != null) && (!IsDeletedItemsFolder(moveToFolder))) ||
                !(item is T) || ItemDeleting == null) return;

            //
            // Listeners to the AppointmentDeleting event can cancel 
            // the move operation if moving to the deleted items folder.
            //
            var args = new CancelEventArgs<T>((T)item);
            ItemDeleting(this, args);
            cancel = args.Cancel;
        }

        private bool IsDeletedItemsFolder(MAPIFolder folder)
        {
            return (folder.EntryID == _deletedItemsFolder.EntryID);
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Releases unmanaged and - optionally - managed resources
        /// </summary>
        /// <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {}

                ((MAPIFolderEvents_12_Event)_folder).BeforeItemMove -= BeforeItemMove;
                _items.ItemChange -= ItemsItemChange;
                _items.ItemAdd -= ItemsItemAdd;

                _items.ReleaseComObject();
                _folder.ReleaseComObject();
                _deletedItemsFolder.ReleaseComObject();
                _items = null;
                _folder = null;
                _deletedItemsFolder = null;
            }
            _disposed = true;
        }
    }
}
