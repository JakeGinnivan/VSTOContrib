using System;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;

namespace VSTOContrib.Core.RibbonFactory
{
    class VstoContribContext
    {
        readonly Type addinType;
        object application;

        public VstoContribContext(Assembly[] assemblies, AddInBase addinBase, string fallbackRibbonType, IViewLocationStrategy viewLocationStrategy = null)
        {
            FallbackRibbonType = fallbackRibbonType;
            Assemblies = assemblies;
            AddinBase = addinBase;
            addinType = addinBase.GetType();

            var globalsType = Type.GetType(addinType.AssemblyQualifiedName.Replace("." + addinType.Name, ".Globals"));
            var factory = globalsType.GetProperty("Factory", BindingFlags.Static | BindingFlags.NonPublic)
                .GetValue(null, null);
            VstoFactory = (Factory)factory;
            ViewLocationStrategy = viewLocationStrategy ?? new DefaultViewLocationStrategy();
            ViewModelFactory = new DefaultViewModelFactory();
            RibbonXmlFromTypeLookup = new Dictionary<string, string>();
            TagToCallbackTargetLookup = new Dictionary<string, CallbackTarget>();
            ErrorHandlers = new List<IErrorHandler>();
        }

        public Assembly[] Assemblies { get; set; }
        public AddInBase AddinBase { get; set; }
        public IViewLocationStrategy ViewLocationStrategy { get; set; }
        public Dictionary<string, string> RibbonXmlFromTypeLookup { get; set; }
        public Dictionary<string, CallbackTarget> TagToCallbackTargetLookup { get; set; }

        public object Application
        {
            get
            {
                if (application == null)
                {
                    const BindingFlags bindingFlags = BindingFlags.Instance | BindingFlags.NonPublic;
                    var fieldInfo = addinType.GetField("Application", bindingFlags);
                    application = fieldInfo.GetValue(AddinBase);
                }
                
                return application;
            }
        }

        public Factory VstoFactory { get; private set; }
        public string FallbackRibbonType { get; private set; }
        public ICollection<IErrorHandler> ErrorHandlers { get; private set; }
        public IViewModelFactory ViewModelFactory { get; set; }
    }
}