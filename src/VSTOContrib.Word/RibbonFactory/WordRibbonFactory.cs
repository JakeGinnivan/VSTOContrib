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
        readonly WordOfficeApplicationEvents wordOfficeApplicationEvents;

        public WordRibbonFactory(AddInBase addinBase, params Assembly[] assemblies)
            :this(new WordOfficeApplicationEvents(), addinBase, UseIfEmpty(assemblies, Assembly.GetCallingAssembly()))
        {
        }

        private WordRibbonFactory(WordOfficeApplicationEvents officeApplicationEvents, AddInBase addinBase, Assembly[] assemblies)
            : base(addinBase, assemblies, new WordViewContextProvider(),
                 officeApplicationEvents, WordRibbonType.WordDocument.GetEnumDescription())
        {
            wordOfficeApplicationEvents = officeApplicationEvents;
        }

        protected override void ShuttingDown()
        {
            wordOfficeApplicationEvents.Dispose();
        }

        protected override void InitialiseRibbonFactoryController(IRibbonFactoryController controller, object application)
        {
            wordOfficeApplicationEvents.Initialise(application);
            wordOfficeApplicationEvents.RegisterOpenDocuments();
        }

        public override void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUi)
        {
            //Word does not raise a new document event when we are starting up, and initialise is too soon
            if (wordOfficeApplicationEvents != null)
                wordOfficeApplicationEvents.RegisterOpenDocuments();
            base.Ribbon_Load(ribbonUi);
        }
    }
}