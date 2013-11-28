using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
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
        private PowerPointViewProvider powerPointViewProvider;

        public PowerPointRibbonFactory(
            AddInBase addInBase,
            params Assembly[] assemblies)
            : base(addInBase, assemblies, new PowerPointViewContextProvider(), PowerPointRibbonType.PowerPointPresentation.GetEnumDescription())
        {
        }

        protected override void ShuttingDown()
        {
            powerPointViewProvider.Initialise();
        }

        protected override void InitialiseRibbonFactoryController(IRibbonFactoryController controller, object application)
        {
            powerPointViewProvider = new PowerPointViewProvider((Application)application);
            controller.Initialise(powerPointViewProvider);
        }

        public override void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUi)
        {
            //Word does not raise a new document event when we are starting up, and initialise is too soon
            powerPointViewProvider.RegisterOpenDocuments();
            base.Ribbon_Load(ribbonUi);
        }
    }
}