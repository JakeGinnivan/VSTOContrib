using System.Collections.Generic;
using VSTOContrib.Core.RibbonFactory.Internal;

namespace VSTOContrib.Core.RibbonFactory
{
    class VstoContribContext<TRibbonTypes> where TRibbonTypes : struct
    {
        public readonly Dictionary<TRibbonTypes, string> RibbonXmlFromTypeLookup;

        /// <summary>
        ///     Lookup from a viewmodel type to it's ribbon XML
        /// </summary>
        public readonly Dictionary<string, CallbackTarget<TRibbonTypes>> TagToCallbackTargetLookup;

        public VstoContribContext()
        {
            RibbonXmlFromTypeLookup = new Dictionary<TRibbonTypes, string>();
            TagToCallbackTargetLookup = new Dictionary<string, CallbackTarget<TRibbonTypes>>();
        }
    }
}