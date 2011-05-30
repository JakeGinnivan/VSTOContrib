using System;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace FacebookToOutlook.Presentation.Behaviours
{
    /// <summary>
    /// Persists a Window's Size, Location and WindowState to UserScopeSettings 
    /// </summary>
    public class WindowSettings
    {
        [DllImport("user32.dll")]
        static extern bool SetWindowPlacement(IntPtr hWnd, [In] ref WINDOWPLACEMENT lpwndpl);

        [DllImport("user32.dll")]
        static extern bool GetWindowPlacement(IntPtr hWnd, out WINDOWPLACEMENT lpwndpl);

        // ReSharper disable InconsistentNaming
        const int SW_SHOWNORMAL = 1;
        const int SW_SHOWMINIMIZED = 2;
        // ReSharper restore InconsistentNaming

        internal class WindowApplicationSettings : ApplicationSettingsBase
        {
            public WindowApplicationSettings(WindowSettings windowSettings)
                : base(windowSettings._window.GetType().FullName) 
            {
            }

            [UserScopedSetting]
            public WINDOWPLACEMENT? Placement
            {
                get
                {
                    if (this["Placement"] != null)
                    {
                        return ((WINDOWPLACEMENT)this["Placement"]);
                    }
                    return null;
                }
                set
                {
                    this["Placement"] = value;
                }
            }
        }
        private Window _window;

        public WindowSettings(Window window)
        {
            _window = window;
        }

        /// <summary>
        /// Register the "Save" attached property and the "OnSaveInvalidated" callback 
        /// </summary>
        public static readonly DependencyProperty SaveProperty
            = DependencyProperty.RegisterAttached("Save", typeof(bool), typeof(WindowSettings),
                                                  new FrameworkPropertyMetadata(new PropertyChangedCallback(OnSaveInvalidated)));

        public static void SetSave(DependencyObject dependencyObject, bool enabled)
        {
            dependencyObject.SetValue(SaveProperty, enabled);
        }

        /// <summary>
        /// Called when Save is changed on an object.
        /// </summary>
        private static void OnSaveInvalidated(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e)
        {
            var window = dependencyObject as Window;
            if (window == null || !((bool) e.NewValue)) return;
            var settings = new WindowSettings(window);
            settings.Attach();
        }

        /// <summary>
        /// Load the Window Size Location and State from the settings object
        /// </summary>
        protected virtual void LoadWindowState()
        {
            Settings.Reload();

            if (Settings.Placement == null) return;
            try
            {
                // Load window placement details for previous application session from application settings
                // if window was closed on a monitor that is now disconnected from the computer,
                // SetWindowPlacement will place the window onto a visible monitor.
                var wp = Settings.Placement.Value;

                wp.length = Marshal.SizeOf(typeof (WINDOWPLACEMENT));
                wp.flags = 0;
                wp.showCmd = (wp.showCmd == SW_SHOWMINIMIZED ? SW_SHOWNORMAL : wp.showCmd);
                var hwnd = new WindowInteropHelper(_window).Handle;
                SetWindowPlacement(hwnd, ref wp);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(string.Format("Failed to load window state:\r\n{0}", ex));
            }
        }

        /// <summary>
        /// Save the Window Size, Location and State to the settings object
        /// </summary>
        protected virtual void SaveWindowState()
        {
            WINDOWPLACEMENT wp;
            var hwnd = new WindowInteropHelper(_window).Handle;
            GetWindowPlacement(hwnd, out wp);
            Settings.Placement = wp;
            Settings.Save();
        }

        private void Attach()
        {
            if (_window == null) return;
            _window.Closing += WindowClosing;
            _window.SourceInitialized += WindowSourceInitialized;
        }

        void WindowSourceInitialized(object sender, EventArgs e)
        {
            LoadWindowState();
        }

        private void WindowClosing(object sender, CancelEventArgs e)
        {
            SaveWindowState();
            _window.Closing -= WindowClosing;
            _window.SourceInitialized -= WindowSourceInitialized;
            _window = null;
        }

        private WindowApplicationSettings _windowApplicationSettings;

        internal virtual WindowApplicationSettings CreateWindowApplicationSettingsInstance()
        {
            return new WindowApplicationSettings(this);
        }

        [Browsable(false)]
        internal WindowApplicationSettings Settings
        {
            get
            {
                if (_windowApplicationSettings == null)
                {
                    _windowApplicationSettings = CreateWindowApplicationSettingsInstance();
                }
                return _windowApplicationSettings;
            }
        }
    }
}