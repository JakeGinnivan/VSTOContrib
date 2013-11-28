using System;

namespace VSTOContrib.Core.RibbonFactory
{
    /// <summary>
    /// Arguments for a new View being Opened
    /// </summary>
    public class NewViewEventArgs<TRibbonTypes> : EventArgs
    {
        readonly object viewInstance;
        readonly object viewContext;
        readonly TRibbonTypes ribbonType;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="viewInstance"></param>
        /// <param name="viewContext"></param>
        /// <param name="ribbonType"></param>
        public NewViewEventArgs(object viewInstance, object viewContext, TRibbonTypes ribbonType)
        {
            this.viewInstance = viewInstance;
            this.viewContext = viewContext;
            this.ribbonType = ribbonType;
        }

        /// <summary>
        /// 
        /// </summary>
        public TRibbonTypes RibbonType
        {
            get { return ribbonType; }
        }

        /// <summary>
        /// 
        /// </summary>
        public object ViewInstance
        {
            get { return viewInstance; }
        }

        ///<summary>
        ///</summary>
        public object ViewContext
        {
            get { return viewContext; }
        }

        /// <summary>
        /// True if a viewmodel was wired up to the view. If false call Marshal.ReleaseComObject on view. 
        /// DO NOT release com object if this property is true
        /// </summary>
        public bool Handled { get; set; }
    }
}