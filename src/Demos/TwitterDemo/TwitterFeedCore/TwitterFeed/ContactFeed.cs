using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Office.Contrib;
using Office.Contrib.Extensions;
using Office.Contrib.RibbonFactory;
using Office.Outlook.Contrib.RibbonFactory;
using TwitterFeedCore.Services;
using CustomTaskPane = Microsoft.Office.Tools.CustomTaskPane;

namespace TwitterFeedCore.TwitterFeed
{
    [RibbonViewModel(OutlookRibbonType.OutlookContact)]
    public class ContactFeed : ViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        private readonly BackgroundWorker _worker = new BackgroundWorker();
        private readonly ITwitterService _twitterService;
        private ContactAdapter _contactAdapter;
        private bool _panelShown;
        private Inspector _inspector;
        private string _username;
        private WpfPanelHost _control;
        private CustomTaskPane _twitterTaskPane;

        public ContactFeed(ITwitterService twitterService)
        {
            _twitterService = twitterService;
            Tweets = new ObservableCollection<Tweet>();
            _worker.DoWork += (sender, e) =>
                                  {
                                      e.Result = _twitterService.GetTwitterStreamForUsername((string) e.Argument);
                                  };
            _worker.RunWorkerCompleted += (sender, e) => System.Windows.Application.Current.Dispatcher.BeginInvoke(
                (System.Action) (() =>
                                     {
                                         foreach (
                                             var tweet in (List<Tweet>) e.Result)
                                             Tweets.Add(tweet);
                                         RaisePropertyChanged("IsBusy");
                                     }));
            RefreshCommand = new RelayCommand(Refresh, CanRefresh);
            SaveUsernameCommand = new RelayCommand(SaveUsername);
        }

        public bool IsBusy { get { return _worker.IsBusy; } }

        private void SaveUsername()
        {
            _contactAdapter.TwitterUsername = TwitterUsername;
            RefreshCommand.Execute(null);
        }

        private void Refresh()
        {
            if (_worker.IsBusy) return;

            _worker.RunWorkerAsync(TwitterUsername);
            RaisePropertyChanged("IsBusy");
        }

        private bool CanRefresh()
        {
            return !string.IsNullOrEmpty(TwitterUsername);
        }

        public void Displayed(object context)
        {
            _inspector = (Inspector) context;
            _contactAdapter = new ContactAdapter((ContactItem) _inspector.CurrentItem);
            TwitterUsername = _contactAdapter.TwitterUsername;
            if (!string.IsNullOrEmpty(TwitterUsername))
                RefreshCommand.Execute(null);
        }

        public string TwitterUsername
        {
            get { return _username; }
            set
            {
                _username = value;
                RaisePropertyChanged("TwitterUsername");
            }
        }

        public ICommand RefreshCommand { get; private set; }

        public ICommand SaveUsernameCommand { get; private set; }

        public IRibbonUI RibbonUi { get; set; }

        public ObservableCollection<Tweet> Tweets { get; private set; }

        public bool PanelShown
        {
            get { return _panelShown; }
            set
            {
                if (_panelShown == value) return;
                _panelShown = value;
                _twitterTaskPane.Visible = value;
                RaisePropertyChanged("PanelShown");
            }
        }

        public void RegisterTaskPanes(Register register)
        {
            _control = new WpfPanelHost
                           {
                               Child = new TwitterFeed
                                           {
                                               DataContext = this
                                           }
                           };

            _twitterTaskPane = register(_control, "Twitter");
            _twitterTaskPane.Visible = true;
            PanelShown = true;
            _twitterTaskPane.VisibleChanged += TwitterTaskPaneVisibleChanged;
            TwitterTaskPaneVisibleChanged(this, EventArgs.Empty);
        }

        public override void Cleanup()
        {
            _twitterTaskPane.VisibleChanged -= TwitterTaskPaneVisibleChanged;
            _contactAdapter.Contact.ReleaseComObject();
            if (_control == null) return;

            _control.Dispose();
            base.Cleanup();
        }

        private void TwitterTaskPaneVisibleChanged(object sender, EventArgs e)
        {
            _panelShown = _twitterTaskPane.Visible;
            RaisePropertyChanged("PanelShown");
        }
    }
}
