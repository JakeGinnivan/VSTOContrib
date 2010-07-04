using System;
using System.ComponentModel;
using System.Reflection;
using Microsoft.Office.Interop.Outlook;

namespace Outlook.Utility
{
    /// <summary>
    /// Provides adapter class for common types of Outlook Items (contacts, appointments etc).
    /// Exposing common properties/methods like EntryId, Save etc.
    /// </summary>
    public class ItemAdapter
    {
        private readonly ItemEvents_10_Event _officeItem;
        readonly Type _outlookItemType;

        /// <summary>
        /// Occurs when the items inspector is opening.
        /// </summary>
        public event CancelEventHandler Opening;
        /// <summary>
        /// Occurs when the item is being saved.
        /// </summary>
        public event CancelEventHandler Writing;
        /// <summary>
        /// Occurs when the items inspector is closing.
        /// </summary>
        public event CancelEventHandler Closing;
        /// <summary>
        /// Occurs when item is being deleted.
        /// </summary>
        public event CancelEventHandler Deleting;

        /// <summary>
        /// Create an adapter from an Outlook Item.
        /// </summary>
        /// <param name="obj">The object to try and adapt.</param>
        /// <returns>Adapter if successful, null if unsuccessful</returns>
        public static ItemAdapter FromObject(object obj)
        {
            if (obj is ItemEvents_10_Event)
                return new ItemAdapter((ItemEvents_10_Event)obj);

            return null;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ItemAdapter"/> class.
        /// </summary>
        /// <param name="item">The item.</param>
        public ItemAdapter(ItemEvents_10_Event item)
        {
            _officeItem = item;
            _outlookItemType = OfficeItem.GetType();

            OfficeItem.Open += OnOpen;
            OfficeItem.Close += OnClose;
            OfficeItem.Write += OnWrite;
            OfficeItem.BeforeDelete += OnDelete;
        }

        void OnClose(ref bool cancel)
        {
            if (Closing == null)
                return;
            var args = new CancelEventArgs(cancel);
            Closing(this, args);
            cancel = args.Cancel;
        }

        void OnDelete(object item, ref bool cancel)
        {
            if (Deleting == null)
                return;
            var args = new CancelEventArgs(cancel);
            Deleting(this, args);
            cancel = args.Cancel;
        }

        void OnWrite(ref bool cancel)
        {
            if (Writing == null)
                return;
            var args = new CancelEventArgs(cancel);
            Writing(this, args);
            cancel = args.Cancel;
        }

        void OnOpen(ref bool cancel)
        {
            if (Opening == null)
                return;
            var args = new CancelEventArgs(cancel);
            Opening(this, args);
            cancel = args.Cancel;
        }

        /// <summary>
        /// Gets the user properties.
        /// </summary>
        /// <value>The user properties.</value>
        public UserProperties UserProperties
        {
            get { return (UserProperties)GetProperty("UserProperties"); }
        }

        /// <summary>
        /// Gets the item properties.
        /// </summary>
        /// <value>The item properties.</value>
        public ItemProperties ItemProperties
        {
            get { return (ItemProperties)GetProperty("ItemProperties"); }
        }

        /// <summary>
        /// Gets the entry id.
        /// </summary>
        /// <value>The entry id.</value>
        public string EntryId
        {
            get { return (string)GetProperty("EntryId"); }
        }

        /// <summary>
        /// Gets the parent.
        /// </summary>
        /// <value>The parent.</value>
        public object Parent
        {
            get { return GetProperty("Parent"); }
        }

        /// <summary>
        /// Gets the class.
        /// </summary>
        /// <value>The class.</value>
        public OlObjectClass Class
        {
            get { return (OlObjectClass)GetProperty("Class"); }
        }

        /// <summary>
        /// Gets the outlook item.
        /// </summary>
        /// <value>The outlook item.</value>
        public ItemEvents_10_Event OfficeItem
        {
            get { return _officeItem; }
        }

        /// <summary>
        /// Saves this instance.
        /// </summary>
        public void Save()
        {
            InvokeMethod("Save");
        }

        /// <summary>
        /// Deletes this instance.
        /// </summary>
        public void Delete()
        {
            InvokeMethod("Delete");
        }

        /// <summary>
        /// Displays this instance.
        /// </summary>
        public void Display()
        {
            InvokeMethod("Display");
        }

        internal object GetProperty(string name)
        {
            return _outlookItemType.InvokeMember(
                name,
                BindingFlags.GetProperty,
                null, OfficeItem, null);
        }

        internal void SetProperty(string name, object value)
        {
            _outlookItemType.InvokeMember(
                name,
                BindingFlags.SetProperty,
                null, OfficeItem, new[] { value });
        }

        internal object InvokeMethod(string name, params object[] parameters)
        {
            return _outlookItemType.InvokeMember(
                name,
                BindingFlags.InvokeMethod,
                null, OfficeItem, parameters);
        }

        /// <summary>
        /// If the adapter adapts a contact, return the <see cref="ContactItem"/>
        /// </summary>
        /// <returns>The contact item</returns>
        public ContactItem GetAsContact()
        {
            if (Class != OlObjectClass.olContact)
                throw new ApplicationException(string.Format("Can't GetAsContact for {0}", Class));

            return (ContactItem)OfficeItem;
        }

        /// <summary>
        /// If the adapter adapts a appointment, return the <see cref="AppointmentItem"/>
        /// </summary>
        /// <returns>The appointment item</returns>
        public AppointmentItem GetAsAppointment()
        {
            if (Class != OlObjectClass.olAppointment)
                throw new ApplicationException(string.Format("Can't GetAsAppointment for {0}", Class));

            return (AppointmentItem)OfficeItem;
        }

        /// <summary>
        /// If the adapter adapts a task, return the <see cref="TaskItem"/>
        /// </summary>
        /// <returns>The task item</returns>
        public TaskItem GetAsTask()
        {
            if (Class != OlObjectClass.olTask)
                throw new ApplicationException(string.Format("Can't GetAsTask for {0}", Class));

            return (TaskItem)OfficeItem;
        }

        /// <summary>
        /// If the adapter adapts a post, return the <see cref="PostItem"/>
        /// </summary>
        /// <returns>The post item</returns>
        public PostItem GetAsPost()
        {
            if (Class != OlObjectClass.olPost)
                throw new ApplicationException(string.Format("Can't GetAsPost for {0}", Class));

            return (PostItem)OfficeItem;
        }

        /// <summary>
        /// If the adapter adapts a mail item, return the <see cref="MailItem"/>
        /// </summary>
        /// <returns>The mail item</returns>
        public MailItem GetAsMail()
        {
            if (Class != OlObjectClass.olMail)
                throw new ApplicationException(string.Format("Can't GetAsMail for {0}", Class));

            return (MailItem)OfficeItem;
        }
    }
}
