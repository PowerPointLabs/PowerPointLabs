using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Interop;

using PPExtraEventHelper;

namespace PowerPointLabs.ColorPicker
{
    /// <summary>
    /// Interaction logic for Magnifier.xaml
    /// </summary>
    public partial class Magnifier : Window
    {
        private const int OVERLAY_OUTLINE_SIZE = 2;

        private MagnifierOverlay magnifierOverlay;
        private System.Windows.Forms.Timer timer;
        private IntPtr hwndMag;
        private List<IntPtr> magFilteredWindows;
        private bool isMagInitialized;
        private float magnificationFactor;
        private Size sourceSize;
        private Size sourceHalfSize;
        private Size actualHalfSize;

        public Magnifier(float magnificationFactor)
        {
            InitializeComponent();

            // Calculate dimensions once
            this.magnificationFactor = magnificationFactor;
            sourceSize.Width = (int)(Width / magnificationFactor);
            sourceSize.Height = (int)(Height / magnificationFactor);
            sourceHalfSize.Width = sourceSize.Width / 2;
            sourceHalfSize.Height = sourceSize.Height / 2;
            actualHalfSize.Width = Width / 2;
            actualHalfSize.Height = Height / 2;

            timer = new System.Windows.Forms.Timer();
            timer.Interval = 100;
            timer.Tick += new EventHandler(Timer_Tick);

            magnifierOverlay = new MagnifierOverlay();
            magnifierOverlay.Loaded += MagnifierOverlay_Loaded;

            Visibility = Visibility.Visible;
        }

        #region Public API
        public new void Show()
        {
            if (isMagInitialized)
            {
                UpdateMagnifier();
                timer.Start();
                magnifierOverlay.Show();
                Visibility = Visibility.Visible;
            }
        }

        public new void Hide()
        {
            timer.Stop();
            magnifierOverlay.Hide();
            Visibility = Visibility.Collapsed;
        }
        #endregion

        #region Handled events

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                if (isMagInitialized = Native.MagInitialize())
                {
                    SetupMagnifier();
                }
            }
            catch
            {
                // Windows XP does not support Magnifier
                TeardownMagnifier();
                isMagInitialized = false;
                return;
            }
        }

        private void Window_Unloaded(object sender, RoutedEventArgs e)
        {
            TeardownMagnifier();
        }

        private void MagnifierOverlay_Loaded(object sender, EventArgs e)
        {
            // Overlay handle is only created after it is loaded
            IntPtr overlayHwnd = new WindowInteropHelper(magnifierOverlay).Handle;
            AddMagnifierFilteredWindow(overlayHwnd);
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
            IntPtr hInst = Native.GetModuleHandle(null);
            IntPtr hwnd = new WindowInteropHelper(this).Handle;

            // Create Magnifier window
            hwndMag = Native.CreateWindowEx(
                (int)Native.ExtendedWindowStyles.WS_EX_LEFT,
                Native.WC_MAGNIFIER, "MagnifierWindow",
                (int)Native.WindowStyles.WS_CHILD |
                (int)Native.WindowStyles.WS_VISIBLE |
                (int)Native.MagnifierStyle.MS_CLIPAROUNDCURSOR,
                0, 0, (int)Width, (int)Height,
                hwnd, IntPtr.Zero, hInst, IntPtr.Zero);

            if (hwndMag == IntPtr.Zero)
            {
                throw new Exception();
            }

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

        private void UpdateMagnifier()
        {
            System.Drawing.Point mousePosition = System.Windows.Forms.Control.MousePosition;
            Native.RECT sourceRect = new Native.RECT();
            sourceRect.Left = mousePosition.X - (int)sourceHalfSize.Width;
            sourceRect.Top = mousePosition.Y - (int)sourceHalfSize.Height;
            sourceRect.Right = sourceRect.Left + (int)sourceSize.Width;
            sourceRect.Bottom = sourceRect.Top + (int)sourceSize.Height;

            // Set the source rectangle for the magnifier control
            Native.MagSetWindowSource(hwndMag, sourceRect);

            // Force redraw
            Native.InvalidateRect(hwndMag, IntPtr.Zero, true);

            Left = mousePosition.X - actualHalfSize.Width;
            Top = mousePosition.Y - actualHalfSize.Height;
            magnifierOverlay.Left = mousePosition.X - magnifierOverlay.HalfSize.X;
            magnifierOverlay.Top = mousePosition.Y - magnifierOverlay.HalfSize.Y;
        }

        private void TeardownMagnifier()
        {
            if (isMagInitialized)
            {
                Native.MagUninitialize();
            }
        }
        #endregion
    }
}
