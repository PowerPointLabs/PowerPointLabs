using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

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
        private Size magnifierHalfSize;
        private Size sourceSize;
        private Size sourceHalfSize;

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
            IntPtr hwndControl;
            IntPtr hwndHost;
            int hostWidth, hostHeight;

            internal const int HOST_ID = 0x00000002;
            internal const string WC_MAGNIFIER = "Magnifier";

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
                                                (IntPtr)HOST_ID,
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

            public IntPtr HwndMagnification
            {
                get { return hwndControl; }
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
            // Calculate dimensions once
            magnifierHalfSize.Width = Width / 2;
            magnifierHalfSize.Height = ActualHeight / 2;
            sourceSize.Width = (int)(Width / magnificationFactor);
            sourceSize.Height = (int)(Height / magnificationFactor);
            sourceHalfSize.Width = sourceSize.Width / 2;
            sourceHalfSize.Height = sourceSize.Height / 2;
            
            magnificationControl = new MagnificationControlHost(Width - OutlineWidth.Width, Height - OutlineWidth.Width);
            MagnificationHostElement.Child = magnificationControl;
            hwndMag = magnificationControl.HwndMagnification;

            if (hwndMag == IntPtr.Zero)
            {
                string errorMsg = "Create MagnifierWindow failed. Insufficient heap or handle table entries.";
                throw new OutOfMemoryException(errorMsg);
            }

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

            // Set the source rectangle for the magnifier control
            Native.RECT sourceRect = new Native.RECT();
            sourceRect.Left = (int)(mousePosition.X - sourceSize.Width);
            sourceRect.Top = (int)(mousePosition.Y - sourceSize.Height);
            sourceRect.Right = sourceRect.Left + (int)sourceSize.Width;
            sourceRect.Bottom = sourceRect.Top + (int)sourceSize.Height;
            Native.MagSetWindowSource(hwndMag, sourceRect);

            // Force redraw
            Native.InvalidateRect(hwndMag, IntPtr.Zero, true);

            // WPF units are affected by monitor's DPI
            mousePosition.X = (int)(mousePosition.X / Utils.Graphics.GetDpiScale());
            mousePosition.Y = (int)(mousePosition.Y / Utils.Graphics.GetDpiScale());

            // Update position
            Left = mousePosition.X - magnifierHalfSize.Width;
            Top = mousePosition.Y - magnifierHalfSize.Height;
        }

        #endregion
    }
}
