using System;
using System.Runtime.InteropServices;

namespace Outlook.Utility
{
    [ComImport, Guid("0002E012-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ICatRegister
    {
        void RegisterCategories(int cCategories, IntPtr rgCategoryInfo);

        void UnRegisterCategories(int cCategories, IntPtr rgcatid);

        void RegisterClassImplCategories(
            [In] ref Guid rclsid,
            int cCategories,
            [In, MarshalAs(UnmanagedType.LPArray)] Guid[] rgcatid);

        void UnRegisterClassImplCategories(
            [In] ref Guid rclsid,
            int cCategories,
            [In, MarshalAs(UnmanagedType.LPArray)] Guid[] rgcatid);

        void RegisterClassReqCategories(
            [In] ref Guid rclsid,
            int cCategories,
            [In, MarshalAs(UnmanagedType.LPArray)] Guid[] rgcatid);

        void UnRegisterClassReqCategories(
            [In] ref Guid rclsid,
            int cCategories,
            [In, MarshalAs(UnmanagedType.LPArray)] Guid[] rgcatid);
    }
}