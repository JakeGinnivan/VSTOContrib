using VSTOContrib.Core.RibbonFactory;

namespace VSTOContrib.Word.RibbonFactory
{
    /// <summary>
    /// Meta data about the Outlook ribbon view model
    /// </summary>
    public class WordRibbonViewModelAttribute : RibbonViewModelAttribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WordRibbonViewModelAttribute"/> class.
        /// </summary>
        public WordRibbonViewModelAttribute()
            : base(WordRibbonType.WordDocument)
        {
        }

        /// <summary>
        /// The type of Inspector or Explorer that the ribbon should be displayed for.
        /// </summary>
        /// <value>The ribbon type.</value>
        public new WordRibbonType Type
        {
            get { return (WordRibbonType)base.Type; }
        }
    }
}