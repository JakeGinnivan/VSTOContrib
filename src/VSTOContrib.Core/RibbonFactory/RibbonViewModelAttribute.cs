using System;
using VSTOContrib.Core.Annotations;

namespace VSTOContrib.Core.RibbonFactory
{
    /// <summary>
    /// Meta data about the ribbon view model
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true, Inherited = true)]
    public class RibbonViewModelAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RibbonViewModelAttribute"/> class.
        /// </summary>
        /// <param name="type">The type.</param>
        public RibbonViewModelAttribute(object type)
        {
            if (type == null) throw new ArgumentNullException("type");
            Type = ((Enum)type).GetEnumDescription();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RibbonViewModelAttribute"/> class.
        /// </summary>
        /// <param name="ribbonType">The ribbon type for example 'Microsoft.Word.Document'</param>
        public RibbonViewModelAttribute(string ribbonType)
        {
            Type = ribbonType;
        }

        /// <summary>
        /// The type of Inspector or Explorer that the ribbon should be displayed for.
        /// </summary>
        /// <value>The ribbon type.</value>
        public string Type { get; private set; }
    }
}
