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
        private readonly PowerPointOfficeApplicationEvents powerPointOfficeApplicationEvents;

        public PowerPointRibbonFactory(AddInBase addInBase, params Assembly[] assemblies)
            :this(new PowerPointOfficeApplicationEvents(), addInBase, UseIfEmpty(assemblies, Assembly.GetCallingAssembly()))
        {
        }

        private PowerPointRibbonFactory(PowerPointOfficeApplicationEvents officeApplicationEvents, AddInBase addInBase, Assembly[] assemblies)
            : base(addInBase, assemblies, new PowerPointViewContextProvider(),
            officeApplicationEvents, PowerPointRibbonType.PowerPointPresentation.GetEnumDescription())
        {
            powerPointOfficeApplicationEvents = officeApplicationEvents;
        }

        protected override void ShuttingDown()
        {
            powerPointOfficeApplicationEvents.Dispose();
        }

        protected override void InitialiseRibbonFactoryController(IRibbonFactoryController controller, object application)
        {
            powerPointOfficeApplicationEvents.Initialise(application);
        }

        public override void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUi)
        {
            //Word does not raise a new document event when we are starting up, and initialise is too soon
            powerPointOfficeApplicationEvents.RegisterOpenDocuments();
            base.Ribbon_Load(ribbonUi);
        }
    }
}