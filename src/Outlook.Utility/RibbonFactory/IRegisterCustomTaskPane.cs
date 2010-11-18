using System.Windows.Forms;
using Microsoft.Office.Tools;

namespace Outlook.Utility.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    public interface IRegisterCustomTaskPane
    {
        /// <summary>
        /// Injection method giving the view model a chance to register task pane(s) with the inspector
        /// </summary>
        /// <param name="register">The register.</param>
        void RegisterTaskPanes(Register register);
    }

    /// <summary>
    /// Allows the registration of custom task pane(s)
    /// </summary>
    public delegate CustomTaskPane Register(UserControl control, string title);
}
