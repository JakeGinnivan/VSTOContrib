using System;
using System.Windows;
using Autofac;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using TwitterFeedOutlookAddin.Core;
using TwitterFeedOutlookAddin.Core.Services;
using VSTOContrib.Autofac;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Outlook.RibbonFactory;

namespace TwitterFeedOutlookAddin
{
    public partial class ThisAddIn
    {
        IContainer container;

        private void ThisAddInStartup(object sender, EventArgs e)
        {
            if (System.Windows.Application.Current == null)
                new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
        }

        void ThisAddInShutdown(object sender, EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
            container.Dispose();
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            var containerBuilder = new ContainerBuilder();

            containerBuilder.RegisterType<TwitterService>().As<ITwitterService>();
            containerBuilder.RegisterRibbonViewModels(typeof(ContactFeed).Assembly);
            container = containerBuilder.Build();
            return new OutlookRibbonFactory(new AutofacViewModelFactory(container), new Lazy<CustomTaskPaneCollection>(() => CustomTaskPanes), Globals.Factory, typeof(ContactFeed).Assembly);
        }

        private void InternalStartup()
        {
            RibbonFactory.Current.SetApplication(Application, this);

            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }
    }
}
