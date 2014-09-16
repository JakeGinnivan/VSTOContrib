using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Interop;

namespace VSTOContrib.Core
{
    /// <summary>
    ///     This class retrieves the IWin32Window from the current active Office window.
    ///     This could be used to set the parent for Windows Forms and MessageBoxes.
    /// </summary>
    /// <example>
    ///     OfficeWin32Window parentWindow = new OfficeWin32Window (ThisAddIn.OutlookApplication.ActiveWindow ());
    ///     MessageBox.Show (parentWindow, "This MessageBox doesn't go behind Outlook !!!", "Attention !", MessageBoxButtons.Ok
    ///     , MessageBoxIcon.Question );
    /// </example>
    public class OfficeWin32Window : IWin32Window
    {
        readonly IntPtr windowHandle = IntPtr.Zero;
        readonly object windowObject;

        public OfficeWin32Window(object windowObject, string lpClassName, string captionSuffix)
        {
            if (windowObject is OfficeWin32Window)
                throw new ArgumentException(
                    "OfficeWin32Window is being wrapped in another instance, this should not happen",
                    "windowObject");
            this.windowObject = windowObject;

            // try to get the HWND ptr from the windowObject / could be an Inspector window or an explorer window
            if (windowObject == null)
            {
                windowHandle = IntPtr.Zero;
            }
            else if (ResolveWindowHandle != null)
            {
                windowHandle = ResolveWindowHandle(windowObject);
            }
            else
            {
                var caption = windowObject
                    .GetType()
                    .InvokeMember("Caption", BindingFlags.GetProperty, null, windowObject, null)
                    .ToString();
                windowHandle = FindWindow(lpClassName, caption + captionSuffix);
            }
        }

        internal static Func<object, IntPtr> ResolveWindowHandle { get; set; }

        public object Window
        {
            get { return windowObject; }
        }

        public IntPtr Handle
        {
            get { return windowHandle; }
        }

        public bool IsClosed
        {
            get { return IsWindow(Handle); }
        }

        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32")]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != GetType()) return false;
            return Equals((OfficeWin32Window) obj);
        }

        public override int GetHashCode()
        {
            return windowHandle.GetHashCode();
        }

        bool Equals(OfficeWin32Window other)
        {
            return windowHandle.Equals(other.windowHandle);
        }
    }
}