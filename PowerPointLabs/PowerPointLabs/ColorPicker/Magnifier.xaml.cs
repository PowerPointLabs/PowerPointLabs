using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;

using PPExtraEventHelper;

namespace PowerPointLabs.ColorPicker
{
    /// <summary>
    /// Interaction logic for Magnifier.xaml
    /// </summary>
    public partial class Magnifier : Window
    {
        private MagnificationControlHost magnificationControl;
        private System.Windows.Forms.Timer timer;

        private IntPtr hwndMag;
        private bool isMagInitialized;
        
        private float magnificationFactor;
        private Size actualHalfSize;
        private Size sourceRectSize;
        private Size sourceRectHalfSize;

        public Magnifier(float magnificationFactor)
        {
            InitializeComponent();

            timer = new System.Windows.Forms.Timer();
            timer.Interval = 100;
            timer.Tick += new EventHandler(Timer_Tick);

            this.magnificationFactor = magnificationFactor;
            Visibility = Visibility.Visible;
        }

        #region Magnification HwndHost Win32 Interop

        private class MagnificationControlHost : HwndHost
        {
            internal const string WC_MAGNIFIER = "Magnifier";

            private IntPtr hwndControl;
            private IntPtr hwndHost;
            private int hostWidth, hostHeight;

            public MagnificationControlHost(double width, double height)
            {
                hostWidth = (int)(width * Utils.Graphics.GetDpiScale());
                hostHeight = (int)(height * Utils.Graphics.GetDpiScale());
            }

            protected override HandleRef BuildWindowCore(HandleRef hwndParent)
            {
                hwndControl = IntPtr.Zero;
                hwndHost = IntPtr.Zero;

                hwndHost = Native.CreateWindowEx((int)Native.ExtendedWindowStyles.WS_EX_LAYERED |
                                                (int)Native.ExtendedWindowStyles.WS_EX_TRANSPARENT, 
                                                "static", "",
                                                (int)Native.WindowStyles.WS_CHILD |
                                                (int)Native.WindowStyles.WS_VISIBLE,
                                                0, 0,
                                                hostWidth, hostHeight,
                                                hwndParent.Handle,
                                                IntPtr.Zero,
                                                IntPtr.Zero,
                                                0);
                Native.SetLayeredWindowAttributes(hwndHost, 0, 255, Native.LayeredWindowAttributeFlags.LWA_ALPHA);

                // Create Magnification control
                hwndControl = Native.CreateWindowEx(0, WC_MAGNIFIER, "MagnificationControl",
                                                (int)Native.WindowStyles.WS_CHILD |
                                                (int)Native.WindowStyles.WS_VISIBLE |
                                                (int)Native.MagnifierStyle.MS_CLIPAROUNDCURSOR,
                                                0, 0,
                                                hostWidth, hostHeight,
                                                hwndHost,
                                                IntPtr.Zero, 
                                                IntPtr.Zero, 
                                                0);

                // Clip window into ellipse
                IntPtr ellipseRegion = Native.CreateRoundRectRgn(0, 0,
                                                hostWidth, hostHeight,
                                                hostWidth, hostHeight);
                Native.SetWindowRgn(hwndHost, ellipseRegion, true);

                // Make window click-through
                int extendedStyle = Native.GetWindowLong(hwndParent.Handle, (int)Native.WindowLong.GWL_EXSTYLE);
                Native.SetWindowLong(hwndParent.Handle, (int)Native.WindowLong.GWL_EXSTYLE, 
                                                extendedStyle | (int)Native.ExtendedWindowStyles.WS_EX_TRANSPARENT);

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
                get { return hwndControl; }
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
            }
        }

        public new void Hide()
        {
            timer.Stop();
            Visibility = Visibility.Collapsed;
        }

        #endregion

        #region Handled events

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                isMagInitialized = Native.MagInitialize();
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
            magnificationControl = new MagnificationControlHost(Width - OutlineWidth.Width, Height - OutlineWidth.Width);
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

        private void TeardownMagnifier()
        {
            if (isMagInitialized)
            {
                Native.MagUninitialize();
            }
        }

        private void UpdateMagnifier()
        {
            System.Drawing.Point mousePosition = System.Windows.Forms.Control.MousePosition;

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
            Left = (mousePosition.X / Utils.Graphics.GetDpiScale()) - actualHalfSize.Width;
            Top = (mousePosition.Y / Utils.Graphics.GetDpiScale()) - actualHalfSize.Height;
        }

        #endregion
    }
}
