using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools;
using VSTOContrib.Core;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.PowerPoint.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    [ComVisible(true)]
    public class PowerPointRibbonFactory : Core.RibbonFactory.RibbonFactory
    {
        private readonly PowerPointViewProvider powerPointViewProvider;

        public PowerPointRibbonFactory(AddInBase addInBase, params Assembly[] assemblies)
            :this(new PowerPointViewProvider(), addInBase, UseIfEmpty(assemblies, Assembly.GetCallingAssembly()))
        {
        }

        private PowerPointRibbonFactory(PowerPointViewProvider viewProvider, AddInBase addInBase, Assembly[] assemblies)
            : base(addInBase, assemblies, new PowerPointViewContextProvider(),
            viewProvider, PowerPointRibbonType.PowerPointPresentation.GetEnumDescription())
        {
            powerPointViewProvider = viewProvider;
        }

        protected override void ShuttingDown()
        {
            powerPointViewProvider.Dispose();
        }

        protected override void InitialiseRibbonFactoryController(IRibbonFactoryController controller, object application)
        {
            powerPointViewProvider.Initialise(application);
        }

        public override void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUi)
        {
            //Word does not raise a new document event when we are starting up, and initialise is too soon
            powerPointViewProvider.RegisterOpenDocuments();
            base.Ribbon_Load(ribbonUi);
        }
    }
}