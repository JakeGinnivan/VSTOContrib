using System.ComponentModel;

namespace TwitterResultsWordAddin
{
    public class MyAddinPanelViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public MyAddinPanelViewModel(int getHashCode)
        {
            Text = getHashCode.ToString();
        }

        public string Text { get; set; }
    }
}
