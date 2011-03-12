using System;
using System.Collections.Generic;

namespace Office.Contrib.RibbonFactory
{
    /// <summary>
    /// 
    /// </summary>
    public class ViewClosedEventArgs : EventArgs
    {
        /// <summary>
        /// All currently open views
        /// </summary>
        public IEnumerable<object> AllViews { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="allViews"></param>
        public ViewClosedEventArgs(IEnumerable<object> allViews)
        {
            AllViews = allViews;
        }
    }
}