using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using UserControl = System.Windows.Forms.UserControl;

namespace VSTOContrib.Core.Wpf
{
    /// <summary>
    /// Windows Forms user control that hosts a WPF user Control.
    /// Includes fix for a redraw issue that is present in Office 2007
    /// </summary>
    [ComVisible(true)]
    [ProgId("Office.Contrib.Wpf.WpfPanelHost")]
    [Guid("59A7E2E3-B0AD-449A-81B9-58398873AC8B")]
    public partial class WpfPanelHost : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WpfPanelHost"/> class.
        /// </summary>
        public WpfPanelHost()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Gets or sets the child WPF user control.
        /// </summary>
        /// <value>The child.</value>
        public UIElement Child
        {
            get
            {
                return elementHost1.Child;
            }
            set
            {
                elementHost1.Child = new ScrollViewer
                {
                    Content = value,
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto
                };
                elementHost1.Width++;
                elementHost1.Dock = DockStyle.Fill;
            }
        }
    }
}
