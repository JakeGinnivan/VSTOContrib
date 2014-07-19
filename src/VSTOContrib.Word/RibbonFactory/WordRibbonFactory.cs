using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Word.RibbonFactory
{
    [ComVisible(true)]
    public class WordRibbonFactory : Core.RibbonFactory.RibbonFactory
    {
        readonly WordViewProvider wordViewProvider;

        public WordRibbonFactory(AddInBase addinBase, params Assembly[] assemblies)
            :this(new WordViewProvider(), addinBase, UseIfEmpty(assemblies, Assembly.GetCallingAssembly()))
        {
        }

        private WordRibbonFactory(WordViewProvider viewProvider, AddInBase addinBase, Assembly[] assemblies)
            : base(addinBase, assemblies, new WordViewContextProvider(),
                 viewProvider, WordRibbonType.WordDocument.GetEnumDescription())
        {
            wordViewProvider = viewProvider;
        }

        protected override void ShuttingDown()
        {
            wordViewProvider.Dispose();
        }

        protected override void InitialiseRibbonFactoryController(IRibbonFactoryController controller, object application)
        {
            wordViewProvider.Initialise(application);
            wordViewProvider.RegisterOpenDocuments();
        }

        public override void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUi)
        {
            //Word does not raise a new document event when we are starting up, and initialise is too soon
            if (wordViewProvider != null)
                wordViewProvider.RegisterOpenDocuments();
            base.Ribbon_Load(ribbonUi);
        }
    }
}