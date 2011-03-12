using Office.Contrib.RibbonFactory;

namespace Office.Word.Contrib.RibbonFactory
{
    /// <summary>
    /// Meta data about the Outlook ribbon view model
    /// </summary>
    public class WordRibbonViewModelAttribute : RibbonViewModelAttribute
    {
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