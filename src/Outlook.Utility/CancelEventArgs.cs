namespace Outlook.Utility
{
    /// <summary>
    /// Cancelable event arg
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class CancelEventArgs<T> : EventArgs<T>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CancelEventArgs&lt;T&gt;"/> class.
        /// </summary>
        /// <param name="aValue">A value.</param>
        public CancelEventArgs(T aValue)
            : base(aValue)
        {
        }

        /// <summary>
        /// Flag specifying if event should be canceled
        /// </summary>
        /// <value><c>true</c> to cancel the event; otherwise, <c>false</c>.</value>
        public bool Cancel { get; set; }
    }
}