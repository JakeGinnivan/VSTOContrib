using System.ComponentModel;

namespace VSTOContrib.Word.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    [DefaultValue(WordDocument)]
    public enum WordRibbonType
    {
        /// <summary>
        /// Word Document Ribbon
        /// </summary>
        [Description("Microsoft.Word.Document")]
        WordDocument = 1
    }
}