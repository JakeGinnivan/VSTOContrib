using System.ComponentModel;

namespace TwitterResultsWordAddin
{
    public class MyAddinPanelViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public MyAddinPanelViewModel()
        {
            Text = "Test";
        }

        public string Text { get; set; }
    }
}
