using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Interop;

namespace VSTOContrib.Core
{
    ///<summary>
    /// This class retrieves the IWin32Window from the current active Office window.
    /// This could be used to set the parent for Windows Forms and MessageBoxes.
    ///</summary>
    ///<example>
    /// OfficeWin32Window parentWindow = new OfficeWin32Window (ThisAddIn.OutlookApplication.ActiveWindow ());   
    /// MessageBox.Show (parentWindow, "This MessageBox doesn't go behind Outlook !!!", "Attention !", MessageBoxButtons.Ok , MessageBoxIcon.Question );
    ///</example>
    public class OfficeWin32Window : IWin32Window
    {
        ///<summary>
        /// The <b>FindWindow</b> method finds a window by it's classname and caption.
        ///</summary>
        ///<param name="lpClassName">The classname of the window (use Spy++)</param>
        ///<param name="lpWindowName">The Caption of the window.</param>
        ///<returns>Returns a valid window handle or 0.</returns>
        [DllImport("user32")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        ///<summary>
        /// This holds the window handle for the found Window.
        ///</summary>
        IntPtr _windowHandle = IntPtr.Zero;

        ///<summary>
        /// The <b>Handle</b> of the Outlook WindowObject.
        ///</summary>
        public IntPtr Handle
        {
            get { return _windowHandle; }
        }

        ///<summary>
        /// The <b>OfficeWin32Window</b> class could be used to get the parent IWin32Window for Windows.Forms and MessageBoxes.
        ///</summary>
        ///<param name="windowObject">The current WindowObject.</param>
        public OfficeWin32Window(object windowObject)
        {
            string caption = windowObject.GetType().InvokeMember("Caption", System.Reflection.BindingFlags.GetProperty, null, windowObject, null).ToString();

            // try to get the HWND ptr from the windowObject / could be an Inspector window or an explorer window
            _windowHandle = FindWindow("rctrl_renwnd32\0", caption);
        }
    }
}
