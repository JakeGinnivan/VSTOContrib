using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using VSTOContrib.Core.Annotations;
using VSTOContrib.Core.RibbonFactory.Internal;

namespace VSTOContrib.Core.RibbonFactory
{
    class RibbonXmlRewriter
    {
        const string OfficeCustomui = "http://schemas.microsoft.com/office/2006/01/customui";
        const string OfficeCustomui4 = "http://schemas.microsoft.com/office/2009/07/customui";

        readonly ControlCallbackLookup controlCallbackLookup;
        readonly VstoContribContext vstoContribContext;
        readonly ViewModelResolver ribbonViewModelResolver;

        public RibbonXmlRewriter(VstoContribContext vstoContribContext, ViewModelResolver ribbonViewModelResolver)
        {
            controlCallbackLookup = new ControlCallbackLookup();
            this.vstoContribContext = vstoContribContext;
            this.ribbonViewModelResolver = ribbonViewModelResolver;
        }

        public void LocateAndRegisterViewXml(
            Type viewModelType, string loadMethodName,
            [CanBeNull] string fallbackRibbonType)
        {
            var resourceText = (string)vstoContribContext.ViewLocationStrategy.GetType()
                .GetMethod("LocateViewForViewModel")
                .MakeGenericMethod(viewModelType)
                .Invoke(vstoContribContext.ViewLocationStrategy, new object[] { });

            XDocument ribbonDoc = XDocument.Parse(resourceText);

            //We have to override the Ribbon_Load event to make sure we get the callback
            XElement customUi =
                ribbonDoc.Descendants(XName.Get("customUI", OfficeCustomui)).SingleOrDefault() ?? 
                ribbonDoc.Descendants(XName.Get("customUI", OfficeCustomui4)).Single();

            customUi.SetAttributeValue("onLoad", loadMethodName);

            //And for automatic image loading support
            if (customUi.Attribute("loadImage") == null)
                customUi.SetAttributeValue("loadImage", "GetPicture");

            foreach (var value in ViewModelRibbonTypesLookupProvider.Instance.GetRibbonTypesFor(viewModelType, fallbackRibbonType))
            {
                WireUpEvents(value, null, ribbonDoc, customUi.GetDefaultNamespace());
                vstoContribContext.RibbonXmlFromTypeLookup.Add(value, ribbonDoc.ToString());
            }
        }

        public string RewriteDynamicXml(string ribbonTypes, string dynamicContext, string dynamicXml)
        {
            var dynamicXmlDoc = XDocument.Parse(dynamicXml);
            
            WireUpEvents(ribbonTypes, dynamicContext, dynamicXmlDoc, dynamicXmlDoc.Root.GetDefaultNamespace());

            return dynamicXmlDoc.ToString();
        }

        void WireUpEvents(string ribbonTypes, string dynamicContext, XContainer ribbonDoc, XNamespace xNamespace)
        {
            //Go through each type of Ribbon 
            foreach (string ribbonControl in controlCallbackLookup.RibbonControls)
            {
                //Get each instance of that control in the ribbon definition file
                IEnumerable<XElement> xElements =
                    ribbonDoc.Descendants(XName.Get(ribbonControl, xNamespace.NamespaceName));

                foreach (XElement xElement in xElements)
                {
                    XAttribute elementId = xElement.Attribute(XName.Get("id"));
                    XAttribute elementQId = xElement.Attribute(XName.Get("idQ"));
                    
                    //Go through each possible callback, Concat with common methods on all controls
                    foreach (string controlCallback in controlCallbackLookup.GetVstoControlCallbacks(ribbonControl))
                    {
                        //Look for a defined callback
                        XAttribute callbackAttribute = xElement.Attribute(XName.Get(controlCallback));

                        if (callbackAttribute == null) continue;
                        if (elementId == null && elementQId == null)
                        {
                            throw new InvalidOperationException(string.Format(
                                "VSTO Contrib Requires controls to have an id or an idQ when callbacks are registered. Control='{0}', Callback='{1}'", 
                                ribbonControl, controlCallback));
                        }

                        string currentCallback = callbackAttribute.Value;
                        //Set the callback value to the callback method defined on this factory
                        string factoryMethodName = controlCallbackLookup.GetFactoryMethodName(ribbonControl,
                            controlCallback);
                        callbackAttribute.SetValue(factoryMethodName);

                        //Set the tag attribute of the element, this is needed to know where to 
                        // direct the callback
                        var id = (elementId ?? elementQId).Value;
                        string callbackTag = BuildTag(ribbonTypes, id, factoryMethodName);
                        vstoContribContext.TagToCallbackTargetLookup.Add(callbackTag, new CallbackTarget(ribbonTypes, dynamicContext, currentCallback));
                        xElement.SetAttributeValue(XName.Get("tag"), (ribbonTypes + id));
                        ribbonViewModelResolver.RegisterCallbackControl(ribbonTypes, currentCallback, id);
                    }
                }
            }
        }

        static string BuildTag(string viewModelType, string elementId, string factoryMethodName)
        {
            return viewModelType + elementId + factoryMethodName;
        }
    }
}