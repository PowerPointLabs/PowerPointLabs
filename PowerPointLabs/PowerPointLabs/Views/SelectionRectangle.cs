using System.Drawing;
using System.Windows.Forms;

namespace PowerPointLabs.Views
{
    public partial class SelectionRectangle : Form
    {
        public SelectionRectangle()
        {
            InitializeComponent();

            ShowInTaskbar = false;
            BackColor = Color.CadetBlue;
            FormBorderStyle = FormBorderStyle.None;

            //SetStyle(ControlStyles.UserPaint, true);
            //SetStyle(ControlStyles.Opaque, true);

            //SizeChanged += (s, e) => Invalidate();
            //Paint += (s, e) =>
            //             {
            //                 e.Graphics.Clear(BackColor);

            //                 using (var pen = new Pen(Color.FromArgb(255, 0, 0, 255)))
            //                 {
            //                     e.Graphics.DrawRectangle(pen, 0, 0, Size.Width - 1, Size.Height - 1);
            //                 }
            //             };
        }

        protected override bool ShowWithoutActivation
        {
            get { return true; }
        }
    }
}
