using System;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Word.RibbonFactory;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        AddinBootstrapper core;

        private void ThisAddInStartup(object sender, EventArgs e)
        {

        }

        private void ThisAddInShutdown(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            core = new AddinBootstrapper();
            WordRibbonFactory.SetApplication(Application);
            RibbonFactory.Current.InitialiseFactory(
                t => (IRibbonViewModel)core.Resolve(t),
                CustomTaskPanes);

            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }
    }

    public class AddinBootstrapper : IDisposable
    {
        public object Resolve(Type type)
        {
            return Activator.CreateInstance(type);
        }

        public T Resolve<T>()
        {
            return (T)Resolve(typeof(T));
        }

        public void Dispose()
        {
        }
    }
}
