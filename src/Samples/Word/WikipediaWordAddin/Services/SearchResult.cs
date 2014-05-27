using System.Diagnostics;
using System.Windows.Input;
using VSTOContrib.Core.Wpf;

namespace WikipediaWordAddin.Services
{
    public class SearchResult
    {
        public string title { get; set; }
        public string snippet { get; set; }
        
        public ICommand OpenLink
        {
            get { return new DelegateCommand(() => Process.Start(string.Format("http://en.wikipedia.org/wiki/{0}", title.Replace(" ", "_")))); }
        }
    }
}