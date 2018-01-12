using System;
using System.ComponentModel;
using System.Configuration;
using System.Windows;

namespace PowerPointLabs.WPF
{
    /// <summary>
    ///   Persists a Window's Size, Location and WindowState to UserScopeSettings
    ///   Taken from: http://stackoverflow.com/a/10271913
    /// </summary>
    public class WindowSettings
    {
        #region Fields

        /// <summary>
        ///   Register the "Save" attached property and the "OnSaveInvalidated" callback
        /// </summary>
        public static readonly DependencyProperty SaveProperty = DependencyProperty.RegisterAttached("Save", typeof(bool), typeof(WindowSettings), new FrameworkPropertyMetadata(OnSaveInvalidated));

        private readonly Window _mWindow;

        private WindowApplicationSettings _mWindowApplicationSettings;

        #endregion Fields

        #region Constructors

        public WindowSettings(Window pWindow) { _mWindow = pWindow; }

        #endregion Constructors

        #region Properties

        [Browsable(false)]
        public WindowApplicationSettings Settings
        {
            get
            {
                if (_mWindowApplicationSettings == null)
                {
                    _mWindowApplicationSettings = CreateWindowApplicationSettingsInstance();
                }
                return _mWindowApplicationSettings;
            }
        }

        #endregion Properties

        #region Methods

        public static void SetSave(DependencyObject pDependencyObject, bool pEnabled) { pDependencyObject.SetValue(SaveProperty, pEnabled); }

        protected virtual WindowApplicationSettings CreateWindowApplicationSettingsInstance() { return new WindowApplicationSettings(this); }

        /// <summary>
        ///   Load the Window Size Location and State from the settings object
        /// </summary>
        protected virtual void LoadWindowState()
        {
            Settings.Reload();
            if (Settings.Location != Rect.Empty)
            {
                _mWindow.Left = Settings.Location.Left;
                _mWindow.Top = Settings.Location.Top;
                _mWindow.Width = Settings.Location.Width;
                _mWindow.Height = Settings.Location.Height;
            }
            else
            {
                // by default, start from bottom right
                _mWindow.Left = SystemParameters.PrimaryScreenWidth - 60 - 100;
                _mWindow.Top = SystemParameters.PrimaryScreenHeight - 60 - 50;
            }

            if (Settings.WindowState != WindowState.Maximized)
            {
                _mWindow.WindowState = Settings.WindowState;
            }
        }

        /// <summary>
        ///   Save the Window Size, Location and State to the settings object
        /// </summary>
        protected virtual void SaveWindowState()
        {
            Settings.WindowState = _mWindow.WindowState;
            Settings.Location = _mWindow.RestoreBounds;
            Settings.Save();
        }

        /// <summary>
        ///   Called when Save is changed on an object.
        /// </summary>
        private static void OnSaveInvalidated(DependencyObject pDependencyObject, DependencyPropertyChangedEventArgs pDependencyPropertyChangedEventArgs)
        {
            var window = pDependencyObject as Window;
            if (window != null)
            {
                if ((bool)pDependencyPropertyChangedEventArgs.NewValue)
                {
                    var settings = new WindowSettings(window);
                    settings.Attach();
                }
            }
        }

        private void Attach()
        {
            if (_mWindow != null)
            {
                _mWindow.Closing += WindowClosing;
                _mWindow.Initialized += WindowInitialized;
                _mWindow.Loaded += WindowLoaded;
            }
        }

        private void WindowClosing(object pSender, CancelEventArgs pCancelEventArgs) { SaveWindowState(); }

        private void WindowInitialized(object pSender, EventArgs pEventArgs) { LoadWindowState(); }

        private void WindowLoaded(object pSender, RoutedEventArgs pRoutedEventArgs)
        {
            if (Settings.WindowState == WindowState.Maximized)
            {
                _mWindow.WindowState = Settings.WindowState;
            }
        }

        #endregion Methods

        #region Nested Types

        public class WindowApplicationSettings : ApplicationSettingsBase
        {
            #region Constructors

            public WindowApplicationSettings(WindowSettings pWindowSettings) { }

            #endregion Constructors

            #region Properties

            [UserScopedSetting]
            public Rect Location
            {
                get
                {
                    if (this["Location"] != null)
                    {
                        return ((Rect)this["Location"]);
                    }

                    return Rect.Empty;
                }
                set { this["Location"] = value; }
            }

            [UserScopedSetting]
            public WindowState WindowState
            {
                get
                {
                    if (this["WindowState"] != null)
                    {
                        return (WindowState)this["WindowState"];
                    }

                    return WindowState.Normal;
                }
                set { this["WindowState"] = value; }
            }

            #endregion Properties
        }

        #endregion Nested Types
    }
}