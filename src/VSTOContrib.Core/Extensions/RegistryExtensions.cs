using System;
using Microsoft.Win32;

namespace VSTOContrib.Core.Extensions
{
    /// <summary>
    /// Helpers for navigating the windows registry
    /// </summary>
    public static class RegistryExtensions
    {
        /// <summary>
        /// Checks if registry key exists.
        /// </summary>
        /// <param name="startingKey">The starting key.</param>
        /// <param name="path">The path.</param>
        /// <returns></returns>
        public static bool Exists(this RegistryKey startingKey, string path)
        {
            return startingKey.OpenSubKey(path) != null;
        }

        /// <summary>
        /// Deletes the registry sub key.
        /// </summary>
        /// <param name="key">The key.</param>
        public static void DeleteKey(this RegistryKey key)
        {
            var startingKey = key.OriginatingHive();
            var path = key.GetPath();

            var lastIndexOf = path.LastIndexOf('\\');
            var keyName = path.Substring(lastIndexOf + 1);
            var parentPath = path.Substring(0, lastIndexOf);
            var parentKey = startingKey.OpenSubKey(parentPath, true);
            parentKey.DeleteSubKeyTree(keyName);
        }

        /// <summary>
        /// Gets the originating hive.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns></returns>
        public static RegistryKey OriginatingHive(this RegistryKey key)
        {
            if (key.Name.StartsWith("HKEY_LOCAL_MACHINE"))
                return Registry.LocalMachine;
            if (key.Name.StartsWith("HKEY_CURRENT_USER"))
                return Registry.CurrentUser;

            throw new InvalidOperationException("Unknown registry hive");
        }

        /// <summary>
        /// Gets the path.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns></returns>
        public static string GetPath(this RegistryKey key)
        {
            var index = key.Name.IndexOf('\\');
            return key.Name.Substring(index + 1);
        }
    }
}
