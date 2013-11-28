using System.ComponentModel;

namespace VSTOContrib.PowerPoint.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    [DefaultValue(PowerPointPresentation)]
    public enum PowerPointRibbonType
    {
        /// <summary>
        /// Word Document Ribbon
        /// </summary>
        [Description("Microsoft.PowerPoint.Presentation")]
        PowerPointPresentation = 1
    }
}