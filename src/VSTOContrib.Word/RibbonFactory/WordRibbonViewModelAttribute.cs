using VSTOContrib.Core.RibbonFactory;

namespace VSTOContrib.Word.RibbonFactory
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

        public WordRibbonViewModelAttribute(string ribbonType) : base(ribbonType)
        {
        }
    }
}