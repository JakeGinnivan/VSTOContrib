using System;

namespace Outlook.Utility.RibbonFactory
{
    /// <summary>
    /// Meta data about the ribbon view model
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class RibbonViewModelAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RibbonViewModelAttribute"/> class.
        /// </summary>
        /// <param name="type">The type.</param>
        public RibbonViewModelAttribute(RibbonType type)
        {
            Type = type;
        }

        /// <summary>
        /// The type of Inspector or Explorer that the ribbon should be displayed for.
        /// </summary>
        /// <value>The ribbon type.</value>
        public RibbonType Type { get; private set; }
    }
}
