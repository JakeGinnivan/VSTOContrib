using System;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using WikipediaWordAddin.Services;

namespace WikipediaWordAddin
{
    public class WikipediaResultsViewModel : NotifyPropertyChanged
    {
        readonly Timer searchTimer = new Timer(500);
        SearchResults searchResults;
        string searchText;
        TaskScheduler uiScheduler;

        public WikipediaResultsViewModel()
        {
            searchTimer.Elapsed += DoSearch;
            Application.Current.Dispatcher.Invoke(new Action(() => uiScheduler = TaskScheduler.FromCurrentSynchronizationContext()));
        }

        void DoSearch(object sender, ElapsedEventArgs e)
        {
            searchTimer.Stop();
            Task.Factory.StartNew(() => new WikipediaService().Search(searchText))
                .ContinueWith(r => SearchResults = r.Result, uiScheduler);
        }

        public void Search(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return;
            searchText = text;
            searchTimer.Start();
        }

        public SearchResults SearchResults
        {
            get { return searchResults; }
            set
            {
                if (Equals(value, searchResults)) return;
                searchResults = value;
                OnPropertyChanged("SearchResults");
            }
        }
    }
}
