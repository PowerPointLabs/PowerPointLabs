using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Log;
using PPExtraEventHelper;

namespace PowerPointLabs.ColorPicker
{
    public partial class MagnifierForm : Form
    {
        private const int OVERLAY_OUTLINE_SIZE = 2;

        private MagnifierOverlay overlay;
        private Timer timer;
        private IntPtr hwndMag;
        private List<IntPtr> magFilteredWindows;
        private bool isMagInitialized;
        private float magnificationFactor;
        private Size sourceSize;
        private Size sourceHalfSize;
        private Size actualHalfSize;

        public MagnifierForm(float magnificationFactor)
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

            // Clip the form into a circle
            GraphicsPath gp = new GraphicsPath();
            gp.AddEllipse(ClientRectangle);
            Region = new Region(gp);
            Visible = false;

            try
            {
                if (isMagInitialized = Native.MagInitialize())
                {
                    SetupMagnifier();
                }
            }
            catch (Exception e)
            {
                // Windows XP does not support Magnifier
                Logger.LogException(e, "MagnifierForm");
                TeardownMagnifier();
                isMagInitialized = false;
                return;
            }

            timer = new Timer();
            timer.Interval = 100;
            timer.Tick += new EventHandler(Timer_Tick);

            overlay = new MagnifierOverlay(Width, Height);
            overlay.Owner = this;
            overlay.Shown += Overlay_Shown;
            FormClosing += MagnifierForm_FormClosing;
        }

        #region Magnifier overlay
        private class MagnifierOverlay : Form
        {
            private Pen outlinePen;
            private GraphicsPath outlinePath;
            private Point positionOffset;

            public MagnifierOverlay(int width, int height)
            {
                TopMost = true;
                ShowInTaskbar = false;
                BackColor = Color.LimeGreen;
                TransparencyKey = BackColor;
                FormBorderStyle = FormBorderStyle.None;
                StartPosition = FormStartPosition.Manual;
                MinimumSize = new Size(1, 1);

                int outlineHalfSize = OVERLAY_OUTLINE_SIZE / 2;
                Width = width + OVERLAY_OUTLINE_SIZE;
                Height = height + OVERLAY_OUTLINE_SIZE;
                positionOffset = new Point(-outlineHalfSize, -outlineHalfSize);

                outlinePen = new Pen(Color.Black, OVERLAY_OUTLINE_SIZE);
                outlinePath = new GraphicsPath();
                outlinePath.AddEllipse(outlineHalfSize, outlineHalfSize, width, height);
            }

            public Point PositionOffset
            {
                get { return positionOffset; }
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.DrawPath(outlinePen, outlinePath);
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
                overlay.Show();
                Visible = true;
            }
        }

        public new void Hide()
        {
            timer.Stop();
            overlay.Hide();
            Visible = false;
        }
        #endregion

        #region Handled events
        private void Overlay_Shown(object sender, EventArgs e)
        {
            // Overlay can only be filtered after it is shown
            AddMagnifierFilteredWindow(overlay.Handle);
        }
        
        private void MagnifierForm_FormClosing(object sender, FormClosingEventArgs e)
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
            IntPtr hInst = Native.GetModuleHandle(null);

            // Create Magnifier window
            hwndMag = Native.CreateWindowEx(
                (int)Native.ExtendedWindowStyles.WS_EX_LEFT, 
                Native.WC_MAGNIFIER, "MagnifierWindow", 
                (int)Native.WindowStyles.WS_CHILD |
                (int)Native.WindowStyles.WS_VISIBLE |
                (int)Native.MagnifierStyle.MS_CLIPAROUNDCURSOR,
                0, 0, Width, Height, 
                Handle, IntPtr.Zero, hInst, IntPtr.Zero);

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

        private void AddMagnifierFilteredWindow(IntPtr handle)
        {
            magFilteredWindows.Add(handle);
            Native.MagSetWindowFilterList(hwndMag, (int)Native.MagnifierFilterMode.MW_FILTERMODE_EXCLUDE,
                                          magFilteredWindows.Count, magFilteredWindows.ToArray());
        }

        private void UpdateMagnifier()
        {
            Native.RECT sourceRect = new Native.RECT();
            sourceRect.Left = MousePosition.X - sourceHalfSize.Width;
            sourceRect.Top = MousePosition.Y - sourceHalfSize.Height;
            sourceRect.Right = sourceRect.Left + sourceSize.Width;
            sourceRect.Bottom = sourceRect.Top + sourceSize.Height;

            // Set the source rectangle for the magnifier control
            Native.MagSetWindowSource(hwndMag, sourceRect);

            // Force redraw
            Native.InvalidateRect(hwndMag, IntPtr.Zero, true);

            Left = MousePosition.X - actualHalfSize.Width;
            Top = MousePosition.Y - actualHalfSize.Height;
            overlay.Left = Left + overlay.PositionOffset.X;
            overlay.Top = Top + overlay.PositionOffset.Y;
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
