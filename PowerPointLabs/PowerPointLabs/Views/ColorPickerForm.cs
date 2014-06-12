using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;

namespace PowerPointLabs.Views
{
    public partial class ColorPickerForm : Form
    {
        [DllImport("Gdi32.dll")]
        public static extern int GetPixel(
        System.IntPtr hdc,    // handle to DC
        int nXPos,  // x-coordinate of pixel
        int nYPos   // y-coordinate of pixel
        );

        [DllImport("User32.dll")]
        public static extern IntPtr GetDC(IntPtr wnd);

        [DllImport("User32.dll")]
        public static extern void ReleaseDC(IntPtr dc);


        private System.Windows.Forms.Panel panel1;
        private System.Timers.Timer timer1;

        public ColorPickerForm()
        {
            SetUp();
            InitializeComponent();
            this.SetStyle(ControlStyles.ResizeRedraw, true);
        }

        public ColorPickerForm(PowerPoint.ShapeRange selectedShapes)
            : this()
        {

        }

        private void ColorPickerForm_Load(object sender, EventArgs e)
        {

        }

        private void SetUp()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.timer1 = new System.Timers.Timer();
            ((System.ComponentModel.ISupportInitialize)(this.timer1)).BeginInit();
            this.SuspendLayout();
            //
            // panel1
            //
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Location = new System.Drawing.Point(216, 8);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(64, 56);
            this.panel1.TabIndex = 0;
            //
            // timer1
            //
            this.timer1.Enabled = true;
            this.timer1.SynchronizingObject = this;
            this.timer1.Elapsed += new System.Timers.ElapsedEventHandler(this.timer1_Elapsed);
            //
            // ColorPickerForm
            //
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(292, 273);
            this.Controls.Add(this.panel1);
            this.Name = "ColorPickerForm";
            this.Text = "ColorPickerForm";
            this.Paint += new System.Windows.Forms.PaintEventHandler(this.ColorPickerForm_Paint);
            ((System.ComponentModel.ISupportInitialize)(this.timer1)).EndInit();
            this.ResumeLayout(false);
        }

        private void ColorPickerForm_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
        {
            Random r = new Random(1);

            for (int x = 0; x < 100; x++)
            {
                SolidBrush b = new SolidBrush(Color.FromArgb(r.Next(255), r.Next(255), r.Next(255)));
                e.Graphics.FillRectangle(b, r.Next(this.ClientSize.Width), r.Next(this.ClientSize.Height), r.Next(100), r.Next(100));
            }
        }

        private void timer1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            Point p = Control.MousePosition;
            IntPtr dc = GetDC(IntPtr.Zero);
            this.panel1.BackColor = ColorTranslator.FromWin32(GetPixel(dc, p.X, p.Y));
            ReleaseDC(dc);
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }
    }
}
