using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace VSTOContrib.Core.RibbonFactory.Internal
{
    internal class ControlCallbackLookup
    {
        /// <summary>
        /// Lookup: Control -> VSTO callback name -> Factory method name
        /// </summary>
        private readonly Dictionary<string, Dictionary<string, string>> _controlCallbackLookup;
        /// <summary>
        /// Lookup: Control -> Factory method name -> VSTO callback name
        /// </summary>
        private readonly Dictionary<string, Dictionary<string, string>> _controlReverseCallbackLookup;

        public ControlCallbackLookup(IDictionary<string, Dictionary<string, Expression<Action<RibbonFactory>>>> controlCallbackLookup)
        {
            _controlCallbackLookup = new Dictionary<string, Dictionary<string, string>>();
            _controlReverseCallbackLookup = new Dictionary<string, Dictionary<string, string>>();

            foreach (var control in controlCallbackLookup.Where(c => c.Key != RibbonFactory.CommonCallbacks))
            {
                _controlCallbackLookup.Add(control.Key, new Dictionary<string, string>());
                _controlReverseCallbackLookup.Add(control.Key, new Dictionary<string, string>());
                foreach (var controlCallbacks in control.Value.Concat(controlCallbackLookup[RibbonFactory.CommonCallbacks]))
                {
                    var methodName = controlCallbacks.Value.GetMethodName();
                    _controlCallbackLookup[control.Key].Add(controlCallbacks.Key, methodName);
                    _controlReverseCallbackLookup[control.Key].Add(methodName, controlCallbacks.Key);
                }
            }
        }

        public IEnumerable<string> RibbonControls
        {
            get { return _controlCallbackLookup.Keys; }
        }

        public IEnumerable<string> GetVstoControlCallbacks(string control)
        {
            //Control specific callbacks + common callbacks
            return _controlCallbackLookup[control].Keys;
        }

        public string GetFactoryMethodName(string control, string vstoCallback)
        {
            return _controlCallbackLookup[control][vstoCallback];
        }

        public string GetVstoCallback(string control, string factoryMethod)
        {
            return _controlReverseCallbackLookup[control][factoryMethod];
        }
    }
}