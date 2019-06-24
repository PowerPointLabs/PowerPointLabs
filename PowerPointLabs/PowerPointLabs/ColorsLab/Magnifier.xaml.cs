using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using PPExtraEventHelper;

namespace PowerPointLabs.ColorsLab
{
    /// <summary>
    /// Interaction logic for Magnifier.xaml
    /// </summary>
    public partial class Magnifier : Window
    {
        private MagnifierOverlay magnifierOverlay;
        private MagnificationControlHost magnificationControl;
        private System.Windows.Threading.DispatcherTimer timer;

        private IntPtr hwndMag;
        private bool isMagInitialized;
        private List<IntPtr> magFilteredWindows;

        private float magnificationFactor;
        private Size actualHalfSize;
        private Size sourceRectSize;
        private Size sourceRectHalfSize;

        public Magnifier(float magnificationFactor)
        {
            InitializeComponent();

            timer = new System.Windows.Threading.DispatcherTimer();
            timer.Interval = new TimeSpan(0, 0, 0, 0, 100);
            timer.Tick += new EventHandler(Timer_Tick);
            
            magnifierOverlay = new MagnifierOverlay();
            magnifierOverlay.Visibility = Visibility.Collapsed;
            magnifierOverlay.Loaded += MagnifierOverlay_Loaded;

            isMagInitialized = false;
            this.magnificationFactor = magnificationFactor;
            Visibility = Visibility.Visible;
        }

        #region Magnification HwndHost Win32 Interop

        private class MagnificationControlHost : HwndHost
        {
            internal const string WC_MAGNIFIER = "Magnifier";

            private IntPtr hwndHost;
            private int hostWidth, hostHeight;

            public MagnificationControlHost(double width, double height)
            {
                hostWidth = (int)(width * Utils.GraphicsUtil.GetDpiScale());
                hostHeight = (int)(height * Utils.GraphicsUtil.GetDpiScale());
            }

            protected override HandleRef BuildWindowCore(HandleRef hwndParent)
            {
                hwndHost = IntPtr.Zero;

                // Make window click-through
                int extendedStyle = Native.GetWindowLong(hwndParent.Handle, (int)Native.WindowLong.GWL_EXSTYLE);
                Native.SetWindowLong(hwndParent.Handle, (int)Native.WindowLong.GWL_EXSTYLE,
                                    extendedStyle |
                                    (int)Native.ExtendedWindowStyles.WS_EX_TRANSPARENT |
                                    (int)Native.ExtendedWindowStyles.WS_EX_LAYERED |
                                    (int)Native.ExtendedWindowStyles.WS_EX_TOOLWINDOW);
                Native.SetWindowLong(hwndParent.Handle, (int)Native.WindowLong.GWL_STYLE, (int)Native.WindowStyles.WS_POPUP);

                // Must be transparent for Magnification to work
                Native.SetLayeredWindowAttributes(hwndParent.Handle, 0, 255, Native.LayeredWindowAttributeFlags.LWA_ALPHA);

                // Create Magnification control host
                hwndHost = Native.CreateWindowEx(0, WC_MAGNIFIER, "MagnificationControl",
                                                (int)Native.WindowStyles.WS_CHILD |
                                                (int)Native.WindowStyles.WS_VISIBLE |
                                                (int)Native.MagnifierStyle.MS_CLIPAROUNDCURSOR,
                                                0, 0,
                                                hostWidth, hostHeight,
                                                hwndParent.Handle,
                                                IntPtr.Zero, 
                                                IntPtr.Zero, 
                                                0);

                // Clip window into ellipse
                IntPtr ellipseRegion = Native.CreateRoundRectRgn(0, 0,
                                                hostWidth, hostHeight,
                                                hostWidth, hostHeight);
                Native.SetWindowRgn(hwndHost, ellipseRegion, true);

                return new HandleRef(this, hwndHost);
            }

            protected override IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
            {
                handled = false;
                return IntPtr.Zero;
            }

            protected override void DestroyWindowCore(HandleRef hwnd)
            {
                Native.DestroyWindow(hwnd.Handle);
            }

            public IntPtr HwndControl
            {
                get { return hwndHost; }
            }

            public int HostWidth
            {
                get { return hostWidth; }
            }

            public int HostHeight
            {
                get { return hostHeight; }
            }
        }

        #endregion

        #region Public API

        public new void Show()
        {
            if (isMagInitialized)
            {
                UpdateMagnifier();
                timer.Start();
                Visibility = Visibility.Visible;
                magnifierOverlay.Visibility = Visibility;
            }
        }

        public new void Hide()
        {
            timer.Stop();
            Visibility = Visibility.Collapsed;
            magnifierOverlay.Visibility = Visibility;
        }

        #endregion

        #region Handled events
        
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // Magnification is disabled on some environments
                if (IsMagnificationApiAvailable())
                {
                    isMagInitialized = Native.MagInitialize();
                }

                if (isMagInitialized)
                {
                    SetupMagnifier();
                }
            }
            catch (Exception exception)
            {
                // Windows XP does not support Magnifier
                Logger.LogException(exception, "Magnifier_Window_Loaded");
                TeardownMagnifier();
                isMagInitialized = false;
            }
            finally
            {
                // Hide after initializing UI
                Hide();
            }
        }

        private void MagnifierOverlay_Loaded(object sender, RoutedEventArgs e)
        {
            IntPtr overlayHwnd = new WindowInteropHelper(magnifierOverlay).Handle;
            AddMagnifierFilteredWindow(overlayHwnd);
        }

        private void Window_Unloaded(object sender, RoutedEventArgs e)
        {
            TeardownMagnifier();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            UpdateMagnifier();
        }

        #endregion

        #region Magnifier methods

        private void SetupMagnifier()
        {
            magFilteredWindows = new List<IntPtr>();

            magnificationControl = new MagnificationControlHost(Width, Height);
            MagnificationHostElement.Child = magnificationControl;
            hwndMag = magnificationControl.HwndControl;

            if (hwndMag == IntPtr.Zero)
            {
                string errorMsg = "Create MagnifierWindow failed. Insufficient heap or handle table entries.";
                throw new OutOfMemoryException(errorMsg);
            }

            // Calculate dimensions once
            actualHalfSize.Width = ActualWidth / 2;
            actualHalfSize.Height = ActualHeight / 2;
            sourceRectSize.Width = magnificationControl.HostWidth / magnificationFactor;
            sourceRectSize.Height = magnificationControl.HostHeight / magnificationFactor;
            sourceRectHalfSize.Width = sourceRectSize.Width / 2;
            sourceRectHalfSize.Height = sourceRectSize.Height / 2;

            // Set the magnification factor
            Native.MAGTRANSFORM magTransformMatrix = new Native.MAGTRANSFORM();
            magTransformMatrix.m00 = magnificationFactor;
            magTransformMatrix.m11 = magnificationFactor;
            magTransformMatrix.m22 = 1.0f;
            Native.MagSetWindowTransform(hwndMag, ref magTransformMatrix);
        }

        private void AddMagnifierFilteredWindow(IntPtr handle)
        {
            magFilteredWindows.Add(handle);
            Native.MagSetWindowFilterList(hwndMag, (int)Native.MagnifierFilterMode.MW_FILTERMODE_EXCLUDE,
                                          magFilteredWindows.Count, magFilteredWindows.ToArray());
        }

        private void TeardownMagnifier()
        {
            if (isMagInitialized)
            {
                Native.MagUninitialize();
            }
        }

        private void UpdateMagnifier()
        {
            System.Drawing.Point mousePosition = WinformUtil.MousePosition;

            // Set the source rectangle for the magnification
            Native.RECT sourceRect = new Native.RECT();
            sourceRect.Left = mousePosition.X - (int)sourceRectHalfSize.Width;
            sourceRect.Top = mousePosition.Y - (int)sourceRectHalfSize.Height;
            sourceRect.Right = sourceRect.Left + (int)sourceRectSize.Width;
            sourceRect.Bottom = sourceRect.Top + (int)sourceRectSize.Height;
            Native.MagSetWindowSource(hwndMag, sourceRect);

            // Force magnification redraw
            Native.InvalidateRect(hwndMag, IntPtr.Zero, true);

            // Update position, WPF units are affected by monitor's DPI
            Left = (mousePosition.X / Utils.GraphicsUtil.GetDpiScale()) - actualHalfSize.Width;
            Top = (mousePosition.Y / Utils.GraphicsUtil.GetDpiScale()) - actualHalfSize.Height;
            magnifierOverlay.Left = Left;
            magnifierOverlay.Top = Top;
        }

        #endregion

        #region Helper methods

        private bool IsMagnificationApiAvailable()
        {
            // Magnification API has a bug with window handle sign-extension
            // on Windows 7 64-bit with 32-bit applications (i.e. Win7 WoW64)
            if (IsOSWindows7() &&
                Environment.Is64BitOperatingSystem &&
                !Environment.Is64BitProcess)
            {
                return false;
            }
            return true;
        }

        private bool IsOSWindows7()
        {
            // Major and minor version
            return Environment.OSVersion.Platform == PlatformID.Win32NT &&
                    Environment.OSVersion.Version.Major == 6 &&
                    Environment.OSVersion.Version.Minor == 1;
        }

        #endregion
    }
}
