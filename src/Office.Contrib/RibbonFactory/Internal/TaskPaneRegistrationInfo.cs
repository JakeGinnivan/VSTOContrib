using System.Windows.Forms;

namespace Office.Contrib.RibbonFactory.Internal
{
    internal class TaskPaneRegistrationInfo
    {
        public UserControl Control { get; set; }
        public string Title { get; set; }

        public TaskPaneRegistrationInfo(UserControl control, string title)
        {
            Control = control;
            Title = title;
        }
    }
}