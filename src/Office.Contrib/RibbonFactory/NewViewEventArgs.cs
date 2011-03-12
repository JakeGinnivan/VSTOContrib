using System;

namespace Office.Contrib.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TRibbonTypes"></typeparam>
    public class NewViewEventArgs<TRibbonTypes> : EventArgs
    {
        private readonly object _viewInstance;
        private readonly TRibbonTypes _ribbonType;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="viewInstance"></param>
        /// <param name="ribbonType"></param>
        public NewViewEventArgs(object viewInstance, TRibbonTypes ribbonType)
        {
            _viewInstance = viewInstance;
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

        /// <summary>
        /// True if a viewmodel was wired up to the view. If false call Marshal.ReleaseComObject on view. 
        /// DO NOT release com object if this property is true
        /// </summary>
        public bool Handled { get; set; }
    }
}