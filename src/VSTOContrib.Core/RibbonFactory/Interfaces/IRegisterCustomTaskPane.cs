using System;
using System.Windows.Forms;
using VSTOContrib.Core.RibbonFactory.Internal;

namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    /// <summary>
    /// Notifies VSTO Contrib that your viewmodel wants to register a custom task pane
    /// </summary>
    public interface IRegisterCustomTaskPane
    {
        /// <summary>
        /// Injection method giving the view model a chance to register task pane(s) with the view
        /// </summary>
        /// <param name="register">The register.</param>
        void RegisterTaskPanes(Register register);
    }

    /// <summary>
    /// Allows the registration of custom task pane(s)
    /// </summary>
    public delegate ICustomTaskPaneWrapper Register(Func<UserControl> controlFactory, string title);
}
