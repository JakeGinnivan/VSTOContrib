using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using TwitterFeedOutlookAddin.Core.Services;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Outlook.RibbonFactory;

namespace TwitterFeedOutlookAddin.Core
{
    [OutlookRibbonViewModel(OutlookRibbonType.OutlookContact)]
    public class ContactFeed : OfficeViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        readonly BackgroundWorker worker = new BackgroundWorker();
        readonly ITwitterService twitterService;
        ContactAdapter contactAdapter;
        bool panelShown;
        string username;
        ICustomTaskPaneWrapper twitterTaskPane;

        public ContactFeed(ITwitterService twitterService)
        {
            this.twitterService = twitterService;
            Tweets = new ObservableCollection<Tweet>();
            worker.DoWork += (sender, e) =>
            {
                e.Result = this.twitterService.GetTwitterStreamForUsername((string) e.Argument);
            };
            worker.RunWorkerCompleted += (sender, e) => System.Windows.Application.Current.Dispatcher.BeginInvoke(
                (System.Action) (() =>
                {
                    foreach (
                        var tweet in (List<Tweet>) e.Result)
                        Tweets.Add(tweet);
                    OnPropertyChanged("IsBusy");
                }));
            RefreshCommand = new DelegateCommand(Refresh, CanRefresh);
            SaveUsernameCommand = new DelegateCommand(SaveUsername);
        }

        public bool IsBusy
        {
            get { return worker.IsBusy; }
        }

        void SaveUsername()
        {
            contactAdapter.TwitterUsername = TwitterUsername;
            RefreshCommand.Execute(null);
        }

        void Refresh()
        {
            if (worker.IsBusy) return;

            worker.RunWorkerAsync(TwitterUsername);
            OnPropertyChanged("IsBusy");
        }

        bool CanRefresh()
        {
            return !string.IsNullOrEmpty(TwitterUsername);
        }

        public Factory VstoFactory { get; set; }

        public void Initialised(object context)
        {
            contactAdapter = new ContactAdapter((ContactItem) context);
            TwitterUsername = contactAdapter.TwitterUsername;
            if (!string.IsNullOrEmpty(TwitterUsername))
                RefreshCommand.Execute(null);
        }

        public void CurrentViewChanged(object currentView)
        {
        }

        public string TwitterUsername
        {
            get { return username; }
            set
            {
                username = value;
                OnPropertyChanged("TwitterUsername");
            }
        }

        public ICommand RefreshCommand { get; private set; }

        public ICommand SaveUsernameCommand { get; private set; }

        public IRibbonUI RibbonUi { get; set; }

        public ObservableCollection<Tweet> Tweets { get; private set; }

        public bool PanelShown
        {
            get { return panelShown; }
            set
            {   
                if (panelShown == value) return;
                panelShown = value;
                twitterTaskPane.Visible = value;
                OnPropertyChanged("PanelShown");
            }
        }

        public void RegisterTaskPanes(Register register)
        {
            twitterTaskPane = register(() => new WpfPanelHost
            {
                Child = new TwitterFeed
                {
                    DataContext = this
                }
            }, "Twitter");
            twitterTaskPane.Visible = true;
            PanelShown = true;
            twitterTaskPane.VisibleChanged += TwitterTaskPaneVisibleChanged;
            TwitterTaskPaneVisibleChanged(this, EventArgs.Empty);
        }

        public void Cleanup()
        {
            twitterTaskPane.VisibleChanged -= TwitterTaskPaneVisibleChanged;
            contactAdapter.Contact.ReleaseComObject();
        }

        void TwitterTaskPaneVisibleChanged(object sender, EventArgs e)
        {
            panelShown = twitterTaskPane.Visible;
            OnPropertyChanged("PanelShown");
        }
    }
}