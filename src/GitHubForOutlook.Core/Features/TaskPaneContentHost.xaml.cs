using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Windows;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.Wpf;

namespace GitHubForOutlook.Core.Features
{
    public partial class TaskPaneContentHost : ITaskPaneContentHost
    {
        ICustomTaskPaneWrapper contentTaskPane;

        public TaskPaneContentHost()
        {
            InitializeComponent();
            DataContext = this;
            ContentItems = new ObservableCollection<ITaskPaneContent>();
            ContentItems.CollectionChanged += ContentItemsOnCollectionChanged;
        }

        void ContentItemsOnCollectionChanged(object sender, NotifyCollectionChangedEventArgs notifyCollectionChangedEventArgs)
        {
            var newVisibleValue = ContentItems.Count > 0;
            if (contentTaskPane != null && newVisibleValue != contentTaskPane.Visible)
            {
                contentTaskPane.Visible = newVisibleValue;
            }
        }

        public UIElement AsUIElement()
        {
            return this;
        }

        public ObservableCollection<ITaskPaneContent> ContentItems { get; private set; }

        public void RegisterSelf(Register register)
        {
            contentTaskPane = register(() => new WpfPanelHost
            {
                Child = this
            }, "GitHub for Outlook", false);
        }

        public void AddOrActivate(ITaskPaneContent taskPaneContent)
        {
            RegisterCloseEvent(taskPaneContent);

            var existingIndex = ContentItems.IndexOf(taskPaneContent);
            if (existingIndex == -1)
                ContentItems.Insert(0, taskPaneContent);
            else
                ContentItems.Move(existingIndex, 0);
        }

        void RegisterCloseEvent(ITaskPaneContent taskPaneContent)
        {
            Action taskPaneContentOnClose = null;
            taskPaneContentOnClose = () =>
            {
                ContentItems.Remove(taskPaneContent);
                taskPaneContent.OnClose -= taskPaneContentOnClose;
            };
            taskPaneContent.OnClose += taskPaneContentOnClose;
        }
    }
}
