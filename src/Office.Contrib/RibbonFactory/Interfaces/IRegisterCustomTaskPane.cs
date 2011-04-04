using System;
using System.Windows.Forms;
using Office.Contrib.RibbonFactory.Internal;

namespace Office.Contrib.RibbonFactory.Interfaces
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
    public delegate ICustomTaskPaneWrapper Register(Func<UserControl> controlFactory, string title);
}
