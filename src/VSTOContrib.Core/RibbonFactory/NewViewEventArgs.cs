using System;

namespace VSTOContrib.Core.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TRibbonTypes"></typeparam>
    public class NewViewEventArgs<TRibbonTypes> : EventArgs
    {
        private readonly object _viewInstance;
        private readonly object _viewContext;
        private readonly TRibbonTypes _ribbonType;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="viewInstance"></param>
        /// <param name="viewContext"></param>
        /// <param name="ribbonType"></param>
        public NewViewEventArgs(object viewInstance, object viewContext, TRibbonTypes ribbonType)
        {
            _viewInstance = viewInstance;
            _viewContext = viewContext;
            _ribbonType = ribbonType;
        }

        /// <summary>
        /// 
        /// </summary>
        public TRibbonTypes RibbonType
        {
            get { return _ribbonType; }
        }

        /// <summary>
        /// 
        /// </summary>
        public object ViewInstance
        {
            get { return _viewInstance; }
        }

        ///<summary>
        ///</summary>
        public object ViewContext
        {
            get { return _viewContext; }
        }

        /// <summary>
        /// True if a viewmodel was wired up to the view. If false call Marshal.ReleaseComObject on view. 
        /// DO NOT release com object if this property is true
        /// </summary>
        public bool Handled { get; set; }
    }
}