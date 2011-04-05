using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using TwitterFeedCore.Services;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Outlook.RibbonFactory;

namespace TwitterFeedCore.TwitterFeed
{
    [RibbonViewModel(OutlookRibbonType.OutlookContact)]
    public class ContactFeed : OfficeViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        private readonly BackgroundWorker _worker = new BackgroundWorker();
        private readonly ITwitterService _twitterService;
        private ContactAdapter _contactAdapter;
        private bool _panelShown;
        private string _username;
        private ICustomTaskPaneWrapper _twitterTaskPane;

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
            RefreshCommand = new DelegateCommand(Refresh, CanRefresh);
            SaveUsernameCommand = new DelegateCommand(SaveUsername);
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

        public void Initialised(object context)
        {
            _contactAdapter = new ContactAdapter((ContactItem)context);
            TwitterUsername = _contactAdapter.TwitterUsername;
            if (!string.IsNullOrEmpty(TwitterUsername))
                RefreshCommand.Execute(null);
        }

        public void CurrentViewChanged(object currentView)
        {
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
            _twitterTaskPane = register(() => new WpfPanelHost
            {
                Child = new TwitterFeed
                {
                    DataContext = this
                }
            }, "Twitter");
            _twitterTaskPane.Visible = true;
            PanelShown = true;
            _twitterTaskPane.VisibleChanged += TwitterTaskPaneVisibleChanged;
            TwitterTaskPaneVisibleChanged(this, EventArgs.Empty);
        }

        public void Cleanup()
        {
            _twitterTaskPane.VisibleChanged -= TwitterTaskPaneVisibleChanged;
            _contactAdapter.Contact.ReleaseComObject();
        }

        private void TwitterTaskPaneVisibleChanged(object sender, EventArgs e)
        {
            _panelShown = _twitterTaskPane.Visible;
            RaisePropertyChanged("PanelShown");
        }
    }
}
